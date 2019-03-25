1# Service
import win32serviceutil
import win32service
import win32event
import win32gui
import servicemanager
import socket

# Get current process
import win32process, win32api, win32con

# Input monitoring
from pynput import mouse, keyboard

# Database
from http.client import HTTPSConnection, HTTPConnection
import json
import ssl

# UI
import PySimpleGUI as sg      
import SysTrayIcon, WindowsBalloonTip

# General
import random
import os
import time, datetime

# Database config
cURI                    = '/postgrest/'

cProductivityTable      = {'name': 'productivity', 		'columns': ['time', 'productivity'] }
cFocusedProgramTable    = {'name': 'focused_program', 	'columns': ['time', 'title', 'program'] } 
cInputTable             = {'name': 'input', 			'columns': ['time', 'mouse_moves', 'mouse_clicks', 'mouse_scrolls', 'keyboard_presses'] } 
cEventsTable      		= {'name': 'events',			'columns': ['time', 'event'] } 

# Client config
cUpdateInterval =   1
cSendInterval   =   5

class ProductivityTracker(object):
	def OnMouseMove(self, x, y):
		self.mMouseMoves += 1
		
	def OnMouseClick(self, x, y, button, pressed):
		if pressed:
			self.mMouseClicks += 1
	
	def OnMouseScroll(self, x, y, dx, dy):
		self.mMouseScrolls += 1

	def OnKeyboardPress(self, key):
		self.mKeyboardPresses += 1

	def OnLock(self):
		print("Locked")
	
	def OnUnlock(self):
		print("Unlocked")


	def start(self):
		# Read Influx credentials
		source_path, _ = os.path.split(__file__)
		f = open( os.path.join(source_path, 'postgrest_cred.json'))
		data = json.load(f)
		self.mAuth = data['token']
		self.mHost = data['host']
		
		# Set up influxDB connection
		self.mConnection = HTTPSConnection(self.mHost, context=ssl._create_unverified_context())

		# Initialize values
		self.mBatches               = {}
		self.mLastBatchSend         = -1
		self.mLastPopup             = -1
		self.mScheduledPopup        = -1
		self.mLastUpdate            = time.time()
		self.mLastInput             = time.time()
		self.mLockedSince           = -1
		self.mMouseMoves            = 0
		self.mMouseClicks           = 0
		self.mMouseScrolls          = 0
		self.mKeyboardPresses       = 0
		self.mTrayIcon              = None
		self.mProductivityWindow    = None

		# Register mouse and keyboard listeners
		self.mMouseListener = mouse.Listener(on_move=self.OnMouseMove, on_click=self.OnMouseClick, on_scroll=self.OnMouseScroll)    # Should we exclude mouse moves? 
		self.mMouseListener.start()
		self.mKeyboardListener = keyboard.Listener(on_press=self.OnKeyboardPress)
		self.mKeyboardListener.start()
		self.mMouseListener.wait()
		self.mKeyboardListener.wait()

		# Set up tray icon
		def icon_stop(inSysTrayIcon):
			self.stop()
		def icon_open(inSysTrayIcon):
			self.OpenProductivityWindow()
		def icon_report_distraction(inSysTrayIcon):
			self.ReportDistraction()
		def icon_report_productivity(inSysTrayIcon):
			self.OpenProductivityWindow()
		
		menu_options = (('Open', None, icon_open),
						('Report Productivity', None, icon_report_productivity),
						('Report Distraction', None, icon_report_distraction))
		self.mTrayIcon = SysTrayIcon.SysTrayIcon('ProductivityTracker.ico', 'Productivity Tracker', menu_options, on_quit=icon_stop, blocking=False)

		# Indicate that we should start running
		self.mIsRunning = True


	def stop(self):
		self.mIsRunning = False
		self.mMouseListener.stop()


	def QueueRow(self, inTable, inRowValues):
		columns = inTable['columns']
		if len(inRowValues) != len(columns):
			win32api.MessageBox(0, f'{inRowValues} does not matches columns {columns}', 'Column mismatch', 0x00000010)

		# Create new batch array if it does not exist yet
		if inTable['name'] not in self.mBatches:
			self.mBatches[inTable['name']] = []

		# Append this row to the batch
		values = {}
		for i in range(0, len(columns)):
			values[columns[i]] = inRowValues[i]
		
		self.mBatches[inTable['name']].append(values)


	def UpdateBatches(self):
		headers = { "Authorization": f"Bearer {self.mAuth}", "Content-Type": "application/json"}

		# Wait for the interval to elapse
		if time.time() - self.mLastBatchSend< cSendInterval:
			return
		
		# Send every batch in the dict of batches
		for table_name, values in self.mBatches.items():
			# Skip empty batches
			if len(values) == 0:
				continue

			body = json.dumps(values)
			try:
				self.mConnection.request('POST', cURI + table_name, body=body, headers=headers)
				response = self.mConnection.getresponse()
			except (ConnectionError, TimeoutError, ConnectionRefusedError, ConnectionAbortedError, ConnectionResetError) as e:
				print(f'Connection error: {e}')
				self.mTrayIcon.show_notification("!!! Retrying connection !!!")

				# Reopen the connection
				self.mConnection.close()
				self.mConnection = HTTPSConnection(self.mHost, context=ssl._create_unverified_context())

				#:TODO: Some kind of back-off strategy			
				
				# And retry next update
				return

			# Check reponse code (201: Created)
			if response.getcode() != 201:
				print(f'Failed to post values, status code: {response.getcode()}. Error code:\n\t{response.read()}')
				return # Will retry next update
			
			response.read() # This is needed to get the connection ready for the next request

			print(f'{len(values)} cells updated for table {table_name}.')

			# Clear batch with 
			self.mBatches[table_name] = []
		
		# Reset last send time
		self.mLastBatchSend = time.time()


	def OpenProductivityWindow(self):
		print("Spawning window")
		# Create layout for window
		layout = [ [sg.Text('How productive do you feel?')], 
					[ sg.Button(f'{x}') for x in range(1,11)]]

		# Show window and read output
		#TODO Wait nonblocking for the result
		self.mProductivityWindow = sg.Window('Productivity Tracker').Layout(layout)        

		# Update last pop up time
		self.mLastPopup = time.time()


	def UpdateProductivityWindow(self):
		if self.mProductivityWindow == None:
			return
		
		form_event, form_values = self.mProductivityWindow.ReadNonBlocking()

		# If event AND values are none, then the window was probably closed,
		# so clean up and show again in a bit
		if form_event == None and form_values == None:
			self.mProductivityWindow.Close()
			self.mProductivityWindow = None
			# :TODO: Reschedule window open
			return
		
		# Nothing happened, just wait until the user does something
		if form_event == None:
			return
		
		# By this point the user has pressed a button

		# Close the form
		self.mProductivityWindow.Close()
		self.mProductivityWindow = None        

		# Write timestamp and productivity to sheet
		values = [ str(datetime.datetime.utcnow()), str(form_event)]
		self.QueueRow(cProductivityTable, values)


	def UpdateProductivityTimer(self):
		cNewDayThreshold    = 6 * 60*60
		cAverageDelay       = 2 * 60*60
		cRandomVariation    = 30*60


		# If we've been locked for longer than 'a day', delete the current popup and defer any popup
		# scheduling until next login
		if self.mLockedSince != -1 and time.time() - self.mLockedSince > cNewDayThreshold:
			if self.mProductivityWindow != None:
				self.mProductivityWindow.Close()
				self.mProductivityWindow = None
			self.mScheduledPopup = -1
			return
			

		# Check if there is a popup scheduled
		if self.mScheduledPopup != -1:
			# Show scheduled popup if it is time
			if self.mScheduledPopup < time.time() and self.mScheduledPopup != -1:
				self.OpenProductivityWindow()
				self.mScheduledPopup = -1
		else:
			# Schedule a new popup
			self.mScheduledPopup = time.time() + cAverageDelay + random.randint(-cRandomVariation, cRandomVariation)

	def ReportDistraction(self):
		# Write timestamp and productivity to sheet
		values = [ str(datetime.datetime.utcnow())]
		self.QueueRow(cEventsTable, values)

		self.mTrayIcon.show_notification('Reported distraction!')


	def UpdateFocusedWindow(self):
		try:
			# Get the process for the currently focused window
			whnd = win32gui.GetForegroundWindow()
			window_text = win32gui.GetWindowText(whnd)
			(_, pid) = win32process.GetWindowThreadProcessId(whnd)
			handle = win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, False, pid)

			# Get the executable name of this process
			filename = win32process.GetModuleFileNameEx(handle, 0)
			_, exe = os.path.split(filename)
		except Exception:
			# We might sometimes select something that does not have a process. Just ignore this event then
			filename = None
			exe = None

		# Write timestamp, title text and exectuble name to sheet
		values = [ str(datetime.datetime.utcnow()), window_text, exe]
		self.QueueRow(cFocusedProgramTable, values)

		# If locked since last update, write the lock time. Reset lock time if any other program has focus
		if exe == "LockApp.exe" and window_text == "Windows Default Lock Screen":
			if self.mLockedSince == -1:
				self.mLockedSince = time.time()
				self.OnLock()
		else:
			# exe is None in the password screen, and due to an alt-tab bug.
			# Having an actual program name definitely means we're logged in 
			if exe != None and self.mLockedSince > 0:
				self.mLockedSince = -1
				self.OnUnlock()


	def UpdateInput(self):
		# Write timestamp, mouse and keyboard input to sheet
		elapsed_time = time.time() - self.mLastUpdate
		values = [    str(datetime.datetime.utcnow()), 
						self.mMouseMoves / elapsed_time, 
						self.mMouseClicks / elapsed_time, 
						self.mMouseScrolls / elapsed_time, 
						self.mKeyboardPresses / elapsed_time]
		self.QueueRow(cInputTable, values)

		# Reset stored valeus
		# Note: We might miss some inputs here because of the timeframe between writting and resetting. 
		#       Theoretically we might also get incorrect values because of thread interrupts. This shouldn't be
		#       common enough to be an issue though. 
		self.mMouseMoves = 0
		self.mMouseClicks = 0
		self.mMouseScrolls = 0
		self.mKeyboardPresses = 0


	def main(self):
		# Show a notification to confirm the main loop has started
		# This is mostly here for debugging purposes
		self.mTrayIcon.show_notification("Started!")
		
		# Main loop
		while self.mIsRunning:
			# Manually pump messages for the tray icon
			win32gui.PumpWaitingMessages()

			# Update the window every frame, so it remains responsive
			self.UpdateProductivityWindow()
			
			# Sleep for a bit, running at roughly 60hz
			time.sleep(0.016)
			
			# Update at a lower interval than the message pump
			if time.time() - self.mLastUpdate > cUpdateInterval:

				self.UpdateInput()
				self.UpdateFocusedWindow()
				self.UpdateProductivityTimer()

				# Send any batches if needed
				self.UpdateBatches()

				# Update timestamp
				self.mLastUpdate = time.time()


if __name__ == '__main__':
	productivity_tracker = ProductivityTracker()
	productivity_tracker.start()
	productivity_tracker.main()
	productivity_tracker.stop()

