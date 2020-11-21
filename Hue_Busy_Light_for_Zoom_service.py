import time
import socket
import win32serviceutil
import servicemanager
import win32event
import win32service

from hbl4z import HueBusyLightForZoom


# pyinstaller --onefile --hiddenimport win32timezone Hue_Busy_Light_for_Zoom_service.py
# Hue_Busy_Light_for_Zoom_service.exe install


class HueBusyLightforZoomServiceWrapper(win32serviceutil.ServiceFramework):
    _svc_name_ = "HueBusyLightForZoom"
    _svc_display_name_ = "Hue Busy Light for Zoom"
    _svc_description_ = "Turns on your Hue light when in a Zoom Meeeting"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)

    def SvcStop(self):
        self.stop()
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        self.start()
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        
        self.main()

    def start(self):
        
        self.hbl4z = HueBusyLightForZoom()

        if self.hbl4z.app_settings['hue_bridge_ip'] == None or self.hbl4z.app_settings['hue_light_id'] == None or self.hbl4z.app_settings['hue_username'] == None:
            raise RuntimeError("Hue settings are missing from registry. please run GUI for initial setup.")

        self.isrunning = True

    def stop(self):
        self.isrunning = False

    def main(self):
        while self.isrunning:
            self.hbl4z.zoom_status_monitor()
            time.sleep(2)


if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(HueBusyLightforZoomServiceWrapper)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(HueBusyLightforZoomServiceWrapper)