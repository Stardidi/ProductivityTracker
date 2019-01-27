# Service
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
1
# Sheets
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools

# UI
import PySimpleGUI as sg      

# General
import random
import os
import time, datetime

# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'

# The ID and range of a sample spreadsheet.
cSpreadSheetID          = '1645s8DXpXflurtXb8OFOe1gFGiDtjiTOiPiTFIRAEBM'
cProductivityTable      = 'Productivity!A1:B'
cFocusedProgramTable    = 'FocusedProgram!A1:C'
cInputTable             = 'Input!A1:C'

cUpdateInterval =   1
cSendInterval   =   5

class AppServerSvc (win32serviceutil.ServiceFramework):
    _svc_name_ = "TestService"
    _svc_display_name_ = "Test Service"

    def __init__(self,args):
        '''
        Constructor of the winservice
        '''
        win32serviceutil.ServiceFramework.__init__(self,args)
        self.hWaitStop = win32event.CreateEvent(None,0,0,None)
        socket.setdefaulttimeout(60)


    def SvcStop(self):
        '''
        Called when the service is asked to stop
        '''
        self.stop()
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        '''
        Called when the service is asked to start
        '''
        self.start()
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_,''))
        self.main()


    def on_mouse_move(self, x, y):
        self.mouse_moves += 1
        
    def on_mouse_click(self, x, y, button, pressed):
        if pressed:
            self.mouse_clicks += 1
    
    def on_mouse_scroll(self, x, y, dx, dy):
        self.mouse_scrolls += 1

    def on_keyboard_press(self, key):
        self.keyboard_presses += 1


    def start(self):
        # Set up spreadsheet service
        store = file.Storage('D:\\Perforce\\David150009\\Y4\\ProductivityTracker\\token.json') #FIXME Find a way to make this file available to the service
        creds = store.get()
        if not creds or creds.invalid:
            flow = client.flow_from_clientsecrets('D:\\Perforce\\David150009\\Y4\\ProductivityTracker\\credentials.json', SCOPES) #FIXME Find a way to make this file available to the service
            creds = tools.run_flow(flow, store)
        service = build('sheets', 'v4', http=creds.authorize(Http()))
        self.spreadsheets = service.spreadsheets()

        # Initialize values
        self.batches = {}
        self.last_batch_send = 0
        self.last_popup = time.time()
        self.last_input_update = time.time()
        self.mouse_moves = 0
        self.mouse_clicks = 0
        self.mouse_scrolls = 0
        self.keyboard_presses = 0

        # Register mouse and keyboard listeners
        self.mouse_listener = mouse.Listener(on_move=self.on_mouse_move, on_click=self.on_mouse_click, on_scroll=self.on_mouse_scroll)    # Should we exclude mouse moves? 
        self.mouse_listener.start()
        self.keyboard_listener = keyboard.Listener(on_press=self.on_keyboard_press)
        self.keyboard_listener.start()
        self.mouse_listener.wait()
        self.keyboard_listener.wait()

        self.isrunning = True


    def stop(self):
        self.isrunning = False
        self.mouse_listener.stop()

    def QueueRow(self, inTable, inRowValues):
        # Create new batch array if it does not exist yet
        if inTable not in self.batches:
            self.batches[inTable] = []

        # Append this row to the batch
        self.batches[inTable].append(inRowValues)

    def UpdateBatches(self):
        # Wait for the interval to elapse
        if time.time() - self.last_batch_send < cSendInterval:
            return
        
        # Send every batch in the dict of batches
        for table, values in self.batches.items():
            # Skip empty batches
            if len(values) == 0:
                continue

            body = {
                "majorDimension": "ROWS",
                "values": values
            }
            result = self.spreadsheets.values().append(spreadsheetId=cSpreadSheetID, range=table, valueInputOption='USER_ENTERED', body=body).execute()
            print('{0} cells updated for table {1}.'.format(result.get('updates').get('updatedCells'), table))

            self.batches[table] = []
        
        # Reset last send time
        self.last_batch_send = time.time()



    def main(self):
        while self.isrunning:
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
            values = [ str(datetime.datetime.now()), window_text, exe]
            self.QueueRow(cFocusedProgramTable, values)


            # Write timestamp, mouse and keyboard input to sheet
            elapsed_time = time.time() - self.last_input_update
            values = [    str(datetime.datetime.now()), 
                            self.mouse_moves / elapsed_time, 
                            self.mouse_clicks / elapsed_time, 
                            self.mouse_scrolls / elapsed_time, 
                            self.keyboard_presses / elapsed_time]
            self.QueueRow(cInputTable, values)
            self.last_input_update = time.time()
            self.mouse_moves = 0
            self.mouse_clicks = 0
            self.mouse_scrolls = 0
            self.keyboard_presses = 0

            if time.time() - self.last_popup > 5:
                print("Spawning window")
                # Create layout for window
                layout = [ [sg.Text('How productive do you feel?')], 
                            [ sg.Button(f'{x}') for x in range(1,11)]]

                # Show window and read output
                #TODO Wait nonblocking for the result
                window = sg.Window('Productivity Tracker').Layout(layout)
                event, values = window.Read()
                window.Close()

                # Write timestamp and productivity to sheet
                values = [ str(datetime.datetime.now()), str(event)]
                self.QueueRow(cProductivityTable, values)

                # Update last pop up time
                self.last_popup = time.time()

            # Send any batches if needed
            self.UpdateBatches()
            
            # Sleep for a bit
            time.sleep(cUpdateInterval)

if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(AppServerSvc)