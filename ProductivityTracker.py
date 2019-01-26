# Service
import win32serviceutil
import win32service
import win32event
import win32gui
import servicemanager
import socket

# Get current process
import win32process, win32api, win32con

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
cSpreadSheetID = '1645s8DXpXflurtXb8OFOe1gFGiDtjiTOiPiTFIRAEBM'
cProductivityTable = 'Productivity!A1:B'
cFocusedProgram = 'FocusedProgram!A1:D'


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




    def start(self):
        self.isrunning = True

        store = file.Storage('D:\\Perforce\\David150009\\Y4\\ProductivityTracker\\token.json') #FIXME Find a way to make this file available to the service
        creds = store.get()
        if not creds or creds.invalid:
            flow = client.flow_from_clientsecrets('D:\\Perforce\\David150009\\Y4\\ProductivityTracker\\credentials.json', SCOPES) #FIXME Find a way to make this file available to the service
            creds = tools.run_flow(flow, store)
        service = build('sheets', 'v4', http=creds.authorize(Http()))
        self.spreadsheets = service.spreadsheets()

    def stop(self):
        self.isrunning = False

    def main(self):
        last_popup = time.time()
        while self.isrunning:
            # Get the process for the currently focused window
            whnd = win32gui.GetForegroundWindow()
            window_text = win32gui.GetWindowText(whnd)
            (_, pid) = win32process.GetWindowThreadProcessId(whnd)
            handle = win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, False, pid)

            # Get the executable name of this process
            filename = win32process.GetModuleFileNameEx(handle, 0)
            _, exe = os.path.split(filename)

            # Write timestamp, title text and exectuble name to sheet
            values = [ [ str(datetime.datetime.now()), window_text, exe] ]
            body = {
                "majorDimension": "ROWS",
                "values": values
            }
            result = self.spreadsheets.values().append(spreadsheetId=cSpreadSheetID, range=cFocusedProgram, valueInputOption='USER_ENTERED', body=body).execute()
            print('{0} cells updated.'.format(result.get('updates').get('updatedCells')))


            if time.time() - last_popup > 5:
                # Create layout for window
                layout = [ [sg.Text('How productive do you feel?')], 
                            [ sg.Button(f'{x}') for x in range(1,11)]]

                # Show window and read output
                #TODO Wait nonblocking for the result
                window = sg.Window('Productivity Tracker').Layout(layout)
                event, values = window.Read()
                window.Close()

                # Write timestamp and productivity to sheet
                values = [ [ str(datetime.datetime.now()), str(event)] ]
                body = {
                    "majorDimension": "ROWS",
                    "values": values
                }
                result = self.spreadsheets.values().append(spreadsheetId=cSpreadSheetID, range=cProductivityTable, valueInputOption='USER_ENTERED', body=body).execute()
                print('{0} cells updated.'.format(result.get('updates').get('updatedCells')))

                # Update last pop up time
                last_popup = time.time()

            # Sleep for a bit
            time.sleep(1)

if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(AppServerSvc)