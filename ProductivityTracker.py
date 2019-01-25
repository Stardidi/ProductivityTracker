import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import random
from pathlib import Path
from time import sleep


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
        while self.isrunning:
            random.seed()
            x = random.randint(1, 100000)
            Path(f'c:\\test\\{x}.txt').touch()
            sleep(5)
        pass

if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(AppServerSvc)