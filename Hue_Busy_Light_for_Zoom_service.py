import time
import socket
import win32serviceutil
import servicemanager
import win32event
import win32service

import winreg
import psutil
import requests

# pyinstaller --onefile --hiddenimport win32timezone Hue_Busy_Light_for_Zoom_service.py
# Hue_Busy_Light_for_Zoom_service.exe install

class HueBusyLightForZoom():
    def __init__(self):
        self.app_settings = {}
        REG_PATH = r"Software\Hue Busy Light for Zoom"

        def get_reg(name):
            try:
                registry_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH, 0,
                                            winreg.KEY_READ)
                value, regtype = winreg.QueryValueEx(registry_key, name)
                winreg.CloseKey(registry_key)
                return value
            except WindowsError:
                return None
        
        self.app_settings['hue_light_id'] = get_reg('hue_light_id')   # Light ID number for the light you want to use as your Zoom busy light       
        self.app_settings['hue_username'] = get_reg('hue_username')   # Hue bridge username, see https://developers.meethue.com/develop/get-started-2/
        self.app_settings['hue_bridge_ip'] = get_reg('hue_bridge_ip') # IP address of bridge, don't know this... goto https://discovery.meethue.com
        self.app_settings['zoom_busy_color'] = [0.67, 0.30] # This is a 'xy' color for the Zoom "red" color
        self.app_settings['zoom_busy_brightness_level'] = 125 # this is the brightness of the Zoom busy light, Hue uses the range 0 - 255

    
    def hue_get_current_light_state(self):
        # get current state of light, so we can revert after
        current_light_response = requests.get(url = f"http://{self.app_settings['hue_bridge_ip']}/api/{self.app_settings['hue_username']}/lights/{self.app_settings['hue_light_id']}")

        # Check if the light is already on
        if current_light_response.json()['state']["on"] == True:
            # Light is already on, see if the color value matches the Zoom busy light
            if current_light_response.json()['state']["xy"] == self.app_settings['zoom_busy_color'] and current_light_response.json()['state']['bri'] == app_settings['zoom_busy_brightness_level']:
                # light is already set to Zoom busy light ( maybe this program crashed? ) so don't save the state
                print("No need to run on light, already set to Zoom busy status")
                return None, True

            # store settings that are common for both white and color bulbs
            hue_light_current_state = {
                "on": current_light_response.json()['state']["on"],
                "bri": current_light_response.json()['state']["bri"],
            }

            # if this is a color bulb, store the color related settings
            if "colormode" in current_light_response.json()["state"]:
                hue_light_current_colormode = current_light_response.json()["state"]["colormode"]

                if hue_light_current_colormode == "xy":
                    hue_light_current_state["xy"] = current_light_response.json()["state"]["xy"]
                elif hue_light_current_colormode == "hs":
                    hue_light_current_state["hue"] = current_light_response.json()["state"]["hue"]
                    hue_light_current_state["sat"] = current_light_response.json()["state"]["sat"]
                elif hue_light_current_colormode == "ct":
                    hue_light_current_state["ct"] = current_light_response.json()["state"]["ct"]
            
            return hue_light_current_state, False
        else:
            return None, False
      

    def zoom_status_monitor(self):

        zoom_meeting_processes = [proc for proc in psutil.process_iter() if proc.name() == "CptHost.exe"]
        
        # if we have atleast one Zoom Meeting process (i.e. we are in a meeting now...)
        if zoom_meeting_processes:
            # check the current state of the Hue light
            hue_light_current_state, busy_light_already_on = self.hue_get_current_light_state()

            # if the light isn't on, turn it on
            if busy_light_already_on == False:
                # turn on the light
                try:
                    response = requests.put(url = f"http://{self.app_settings['hue_bridge_ip']}/api/{self.app_settings['hue_username']}/lights/{self.app_settings['hue_light_id']}/state", json={"on":True, "xy":self.app_settings['zoom_busy_color'] , "bri": self.app_settings['zoom_busy_brightness_level']})
                except Exception as e:
                    print(f"Unable to turn on light due to error {e}")

                if response.status_code==200:
                    print("Successfully turned on busy light.")

            # Wait here until we are no longer in the Meeting
            gone = psutil.wait_procs(zoom_meeting_processes)
            
            # Now that wait_procs is over, that means we are not in a Meeting anymore. So let's turn off (or revert) the Hue light
            if hue_light_current_state:
                # we have a previous state for the light, so let's change the light back to the previous color before Zoom Busy light turned on..
                try:
                    response = requests.put(url = f"http://{self.app_settings['hue_bridge_ip']}/api/{self.app_settings['hue_username']}/lights/{self.app_settings['hue_light_id']}/state", json=hue_light_current_state)
                except Exception as e:
                    logging.error(f"Unable to turn off light due to error {e}")
            else:
                # light was not on before Zoom busy light, so just turn off
                try:
                    response = requests.put(url = f"http://{self.app_settings['hue_bridge_ip']}/api/{self.app_settings['hue_username']}/lights/{self.app_settings['hue_light_id']}/state", json={"on":False})
                except Exception as e:
                    logging.error(f"Unable to turn off light due to error {e}")
            
            if response.status_code==200:
                print("Successfully turned off busy light.")

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

        if self.hbl4z.app_settings['hue_bridge_ip'] == "" or self.hbl4z.app_settings['hue_bridge_ip'] == None or self.hbl4z.app_settings['hue_light_id'] == "" or self.hbl4z.app_settings['hue_light_id'] == None or self.hbl4z.app_settings['hue_username'] == "" or self.hbl4z.app_settings['hue_username'] == None:
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