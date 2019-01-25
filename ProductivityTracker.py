import win32serviceutil
import win32service
import win32event
import win32gui
import win32process, win32api, win32con
import servicemanager
import socket
import random
import os
import time


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

    def stop(self):
        self.isrunning = False

    def main(self):
        f = open(f'c:\\test\\active_window.txt', 'w', encoding='utf-8')
        while self.isrunning:
            # Get the process for the currently focused window
            whnd = win32gui.GetForegroundWindow()
            window_text = win32gui.GetWindowText(whnd)
            (_, pid) = win32process.GetWindowThreadProcessId(whnd)
            handle = win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, False, pid)

            # Get the executable name of this process
            filename = win32process.GetModuleFileNameEx(handle, 0)
            _, exe = os.path.split(filename)

            # Write timestamp, title text, and executable name
            current_time = time.time()
            f.write(f'{current_time};{window_text};{exe}\n')
            f.flush()

            # Sleep for a bit
            time.sleep(1)
        pass
        f.close()

if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(AppServerSvc)