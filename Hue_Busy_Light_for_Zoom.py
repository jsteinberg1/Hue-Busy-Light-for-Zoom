import os, sys, time, datetime
import threading
import tkinter as tk
import tkinter.scrolledtext as scrolledtext

import requests
import psutil

# to hide the console you need to add the --noconsole to your pyinstaller command.

root = tk.Tk()

app_settings = {}
connected_to_bridge = False

class PrintLogger:  # create file like object
    def __init__(self, textbox):  # pass reference to text widget
        self.textbox = textbox  # keep ref

    def write(self, text):
        self.textbox.insert(tk.END, str(text))  # write text to textbox
        # could also scroll to end of textbox here to make sure always visible

    def flush(self):  # needed for file like object
        pass


class ZoomProcessMonitorWorker(threading.Thread):
    def run(self):
        # long process goes here
        global stop_workers
        
        while not stop_workers:
            zoom_status_monitor()

class BridgeWorker(threading.Thread):
    def run(self):
        # long process goes here
        global app_settings
        global connected_to_bridge

        while not stop_workers and not connected_to_bridge:
            connect_to_bridge()
            time.sleep(5)
    


def draw_hue_light_list(hue_light_list: list, hue_selected_light_name: str):
    global entry_light_name_text_entry

    label_light_name = tk.Label(root, text="Select a light")

    entry_light_name_text_entry = tk.StringVar()
    entry_light_id = tk.OptionMenu(root, entry_light_name_text_entry, *hue_light_list)
    entry_light_name_text_entry.set(hue_selected_light_name)

    label_light_name.grid(row=1, column=0)
    entry_light_id.grid(row=1, column=1)


def draw_bridge_ip(hue_bridge_ip: str):
    global hue_bridge_ip_text_entry

    # Bridge IP Label
    label_bridge_ip = tk.Label(root, text="Hue Bridge IP Address")  

    hue_bridge_ip_text_entry = tk.StringVar()
    entry_bridge_ip = tk.Entry(root, textvariable=hue_bridge_ip_text_entry)

    label_bridge_ip.grid(row=0, column=0)
    entry_bridge_ip.grid(row=0, column=1)

    hue_bridge_ip_text_entry.set(hue_bridge_ip)


def draw_main_gui():
    
    root.title("Hue Busy Light for Zoom")

    read_settings()

    # Save button
    button = tk.Button(
        text="Save Changes",
        command=lambda: processUserChange(hue_bridge_ip_text_entry.get(), entry_light_name_text_entry.get()),
    )

    # Display console logging data inside tkinter
    console_logging_window = scrolledtext.ScrolledText(root, undo=True)
    console_output_window = PrintLogger(console_logging_window)
    sys.stdout = console_output_window

    # Layout grid
    button.grid(row=2, columnspan=2)
    console_logging_window.grid(row=3, columnspan=2)

    # Start Workers
    global stop_workers
    stop_workers = False

    zw = ZoomProcessMonitorWorker()
    zw.start()

    bw = BridgeWorker()
    bw.start()

    root.mainloop()
    
    stop_workers = True
    
    zw.join()
    bw.join()
    

def save_setting(name, value):

    if os.name == "nt":
        # Write application settings to Windows registry
        import winreg
        REG_PATH = r"Software\Hue Busy Light for Zoom"

        try:
            winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_PATH)
            registry_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH, 0, 
                                        winreg.KEY_WRITE)
            winreg.SetValueEx(registry_key, name, 0, winreg.REG_SZ, value)
            winreg.CloseKey(registry_key)
            return True
        except WindowsError:
            return False


def read_settings():
    # Load application settings
    global app_settings

    if os.name == "nt":
        # Read from windows registry
        import winreg
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
        
        # Please update these three variables with your Philip Hue configuration
        
        app_settings['hue_light_id'] = get_reg('hue_light_id')   # Light ID number for the light you want to use as your Zoom busy light       
        app_settings['hue_username'] = get_reg('hue_username')   # Hue bridge username, see https://developers.meethue.com/develop/get-started-2/
        app_settings['hue_bridge_ip'] = get_reg('hue_bridge_ip') # IP address of bridge, don't know this... goto https://discovery.meethue.com
        
    # TODO MAC OSX

    app_settings['zoom_busy_color'] = [0.67, 0.30] # This is a 'xy' color for the Zoom "red" color
    app_settings['zoom_busy_brightness_level'] = 125 # this is the brightness of the Zoom busy light, Hue uses the range 0 - 255

    return app_settings

   
def connect_to_bridge():
    global app_settings
    global connected_to_bridge
    
    # if we don't have the bridge IP in registry, try to find from Hue discovery
    if app_settings['hue_bridge_ip'] == None:
        print("Performing initial Bridge Discovery, please disconnect from any VPN.")
        # Attempt to find Hue Bridge IP using Philips Hue Discovery method

        try:
            response = requests.get("https://discovery.meethue.com")
            hue_discovery_result = response.json()

            if len(hue_discovery_result) > 0:
                app_settings['hue_bridge_ip'] = hue_discovery_result[0]['internalipaddress']
                print(f"Found Hue Bridge at IP {app_settings['hue_bridge_ip']}...")

            # Save Bridge IP
            save_setting("hue_bridge_ip", app_settings['hue_bridge_ip'])
        except:
            print("Unable to automatically discover Hue Bridge IP. Please enter manually...")

    draw_bridge_ip(app_settings['hue_bridge_ip'])

    # Authenticate to bridge
    if app_settings['hue_bridge_ip'] != None and app_settings['hue_username'] == None:
        # we don't have a username in registry, so prompt user to enroll using link button
        try:
            response = requests.post(f"http://{app_settings['hue_bridge_ip']}/api",json={"devicetype":"hue_busy_lamp_zoom"})
            api_response = response.json()
            app_settings['hue_username'] = api_response[0]['success']['username']
        except:

            error_description = ""
            error_type = ""
            try:
                error_description = api_response[0]['error']['description']
                error_type = api_response[0]['error']['type']
            except:
                print("Unable to connect to bridge...")

            if error_type == 101 or error_description == "link button not pressed":
                print("Press the link button on the Hue Bridge to authorize this application.")
        else:
            save_setting("hue_username", app_settings['hue_username'])
    
    # Light ID

    if app_settings['hue_bridge_ip'] != None and app_settings['hue_username'] != None:
        print("Searching for lights using Hue Bridge API....")

        try:
            light_get_response = requests.get(url = f"http://{app_settings['hue_bridge_ip']}/api/{app_settings['hue_username']}/lights")
            light_response_json = light_get_response.json()
        except:
            print("Unable to Connect to Bridge")
        else:
            
            hue_light_dict = {}

            if len(light_response_json) > 0:
                for light in light_response_json:
                    if "colormode" in light_response_json[light]["state"]:
                        hue_light_dict[light] = light_response_json[light]['name']
            
            hue_light_list = [f"{light_response_json[light]['name']} (id {light})" for light in light_response_json if "colormode" in light_response_json[light]["state"]]
            
            if light_get_response.status_code==200:
                print(f"Successfully connected to bridge, found {len(hue_light_list)} color capable bulbs.")
                connected_to_bridge = True
            else:
                print("Error: Unable to retrieve list of bulbs from bridge...")

            if app_settings['hue_light_id'] == "" or app_settings['hue_light_id'] == None:
                hue_selected_light_name = None
                print("Please select a light above and click Save.")
            else:
                hue_selected_light_name = f"{hue_light_dict[app_settings['hue_light_id']]} (id {app_settings['hue_light_id']})"
            
            draw_hue_light_list(hue_light_list,hue_selected_light_name)



    return app_settings


def processUserChange(hue_bridge_ip, entry_light_name):
    global app_settings
    app_settings['hue_bridge_ip'] = hue_bridge_ip
    save_setting("hue_bridge_ip",hue_bridge_ip)
    app_settings['hue_light_id'] = entry_light_name[:-1][entry_light_name.rfind("(id ") + 4:]
    save_setting("hue_light_id",app_settings['hue_light_id'])


def hue_get_current_light_state():
    # get current state of light, so we can revert after
    current_light_response = requests.get(url = f"http://{app_settings['hue_bridge_ip']}/api/{app_settings['hue_username']}/lights/{app_settings['hue_light_id']}")

    # Check if the light is already on
    if current_light_response.json()['state']["on"] == True:
        # Light is already on, see if the color value matches the Zoom busy light
        if current_light_response.json()['state']["xy"] == app_settings['zoom_busy_color'] and current_light_response.json()['state']['bri'] == app_settings['zoom_busy_brightness_level']:
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


def turn_on_hue_zoom_busy_light(zoom_meeting_processes):
    # turn on the light
    try:
        response = requests.put(url = f"http://{app_settings['hue_bridge_ip']}/api/{app_settings['hue_username']}/lights/{app_settings['hue_light_id']}/state", json={"on":True, "xy":app_settings['zoom_busy_color'] , "bri": app_settings['zoom_busy_brightness_level']})
    except Exception as e:
        print(f"Unable to turn on light due to error {e}")

    if response.status_code==200:
        print("Successfully turned on busy light.")
    

def turn_off_hue_zoom_busy_light(gone):
    
    if gone:
        print("Zoom Meeting Window has been closed...")


def zoom_status_monitor():
    global app_settings

    if app_settings['hue_bridge_ip'] == "" or app_settings['hue_bridge_ip'] == None or app_settings['hue_light_id'] == "" or app_settings['hue_light_id'] == None or app_settings['hue_username'] == "" or app_settings['hue_username'] == None:
        # we don't have the settings to connect to the bridge, sleep and retry
        time.sleep(5)
        return

    if os.name == "nt":
        # Windows
        zoom_meeting_processes = [proc for proc in psutil.process_iter() if proc.name() == "CptHost.exe"]
    
    # TODO add OSX here...

    # Log all Zoom meeting processes
    for process in zoom_meeting_processes:
        print("Detected Zoom Meeting window...")

    # if we have atleast one Zoom Meeting process (i.e. we are in a meeting now...)
    if zoom_meeting_processes:
        # check the current state of the Hue light
        hue_light_current_state, busy_light_already_on = hue_get_current_light_state()

        # if the light isn't on, turn it on
        if busy_light_already_on == False:
            turn_on_hue_zoom_busy_light(zoom_meeting_processes)

        # Wait here until we are no longer in the Meeting (leverage psutil.wait_procs )
        gone = psutil.wait_procs(zoom_meeting_processes, callback=turn_off_hue_zoom_busy_light)
        
        # Now that wait_procs is over, that means we are not in a Meeting anymore. So let's turn off (or revert) the Hue light
        if hue_light_current_state:
            # we have a previous state for the light, so let's change the light back to the previous color before Zoom Busy light turned on..
            try:
                response = requests.put(url = f"http://{app_settings['hue_bridge_ip']}/api/{app_settings['hue_username']}/lights/{app_settings['hue_light_id']}/state", json=hue_light_current_state)
            except Exception as e:
                logging.error(f"Unable to turn off light due to error {e}")
        else:
            # light was not on before Zoom busy light, so just turn off
            try:
                response = requests.put(url = f"http://{app_settings['hue_bridge_ip']}/api/{app_settings['hue_username']}/lights/{app_settings['hue_light_id']}/state", json={"on":False})
            except Exception as e:
                logging.error(f"Unable to turn off light due to error {e}")
        
        if response.status_code==200:
            print("Successfully turned off busy light.")

    else:
        # not in a meeting, wait 2 seconds and start over
        time.sleep(2)


if __name__ == "__main__":
    draw_main_gui()
    #root.protocol("WM_DELETE_WINDOW", root.iconify)
    
    