import winreg
import psutil
import requests

# TODO
'''
1) If the light has changed from the Zoom busy light to something else (e.g. another color), then don't turn off the light when meeting ends....
'''

class HueBusyLightForZoom():
    def __init__(self):
        self.read_settings()

    def read_settings(self):
        self.app_settings = {}
        self.app_settings['REG_PATH'] = r"Software\Hue Busy Light for Zoom"

        def get_reg(name):
            try:
                registry_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.app_settings['REG_PATH'], 0,
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


    def save_setting(self, name, value):

        try:
            winreg.CreateKey(winreg.HKEY_CURRENT_USER, self.app_settings['REG_PATH'])
            registry_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.app_settings['REG_PATH'], 0, 
                                        winreg.KEY_WRITE)
            winreg.SetValueEx(registry_key, name, 0, winreg.REG_SZ, value)
            winreg.CloseKey(registry_key)
            
        except WindowsError:
            return False
        else:
            self.app_settings[name] = value
            return True

        
    def hue_get_current_light_state(self):
        # get current state of light, so we can revert after
        current_light_response = requests.get(url = f"http://{self.app_settings['hue_bridge_ip']}/api/{self.app_settings['hue_username']}/lights/{self.app_settings['hue_light_id']}")

        # Check if the light is already on
        if current_light_response.json()['state']["on"] == True:
            # Light is already on, see if the color value matches the Zoom busy light
            if current_light_response.json()['state']["xy"] == self.app_settings['zoom_busy_color'] and current_light_response.json()['state']['bri'] == self.app_settings['zoom_busy_brightness_level']:
                # light is already set to Zoom busy light ( maybe this program crashed? ) so don't save the state
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
            try:
                hue_light_previous_state, busy_light_already_on = self.hue_get_current_light_state()
            except Exception as e:
                print(f"Unable to get current light status {e}")
                return

            # if the light isn't on, turn it on
            if busy_light_already_on == False:
                # turn on the light
                try:
                    response = requests.put(url = f"http://{self.app_settings['hue_bridge_ip']}/api/{self.app_settings['hue_username']}/lights/{self.app_settings['hue_light_id']}/state", json={"on":True, "xy":self.app_settings['zoom_busy_color'] , "bri": self.app_settings['zoom_busy_brightness_level']})
                except Exception as e:
                    print(f"Unable to turn on light due to error {e}")
                    return

                if response.status_code==200:
                    print("Successfully turned on busy light.")
                else:
                    return
            else:
                print("Busy light is already on, no need to turn it on again...")

            # Wait here until we are no longer in the Meeting
            gone = psutil.wait_procs(zoom_meeting_processes)
            
            # Now that wait_procs is over, that means we are not in a Meeting anymore. So let's turn off (or revert) the Busy Light

            try:
                hue_light_current_state, busy_light_already_on = self.hue_get_current_light_state()
            except Exception as e:
                print(f"Unable to get current light status3 {e}")
                return
                        
            if hue_light_previous_state and busy_light_already_on:
                # we have a previous state for the light, so let's change the light back to the previous color before Zoom Busy light turned on..
                json_payload_for_hue_light_change = hue_light_previous_state

            elif busy_light_already_on:               
                # light was not on before Zoom busy light, so just turn off
                json_payload_for_hue_light_change = {"on":False}
            
            else:
                print("Busy light is already off, no need to turn it off..")
                return
            
            # apply change to Hue Bridge
            try:
                response = requests.put(url = f"http://{self.app_settings['hue_bridge_ip']}/api/{self.app_settings['hue_username']}/lights/{self.app_settings['hue_light_id']}/state", json=json_payload_for_hue_light_change)
            except Exception as e:
                logging.error(f"Unable to apply light command {json_payload_for_hue_light_change} due to error {e}")
            else:
                print("Successfully turned off busy light.")
            

            
