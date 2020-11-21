import os, sys, time, datetime
import threading
import tkinter as tk
import tkinter.scrolledtext as scrolledtext

import requests

from hbl4z import HueBusyLightForZoom

# to hide the console you need to add the --noconsole to your pyinstaller command.


class PrintLogger:  # create file like object
    def __init__(self, textbox):  # pass reference to text widget
        self.textbox = textbox  # keep ref

    def write(self, text):
        self.textbox.insert(tk.END, str(text))  # write text to textbox
        # could also scroll to end of textbox here to make sure always visible

    def flush(self):  # needed for file like object
        pass

class tk_hue_light_for_zoom_app():
    def __init__(self):
        self.root = tk.Tk()
        self.connected_to_bridge = False
        self.entry_light_name_text_entry = ""
        self.hue_bridge_ip_text_entry = ""
        self.stop_workers = False

        self.hbl4z = HueBusyLightForZoom()

        self.tk_mainloop()

    def draw_hue_light_list(self, hue_light_list: list, hue_selected_light_name: str):
        global entry_light_name_text_entry

        label_light_name = tk.Label(self.root, text="Select a light")

        self.entry_light_name_text_entry = tk.StringVar()
        entry_light_id = tk.OptionMenu(self.root, self.entry_light_name_text_entry, *hue_light_list)
        self.entry_light_name_text_entry.set(hue_selected_light_name)

        label_light_name.grid(row=1, column=0)
        entry_light_id.grid(row=1, column=1)


    def draw_bridge_ip(self, hue_bridge_ip: str):
        # Bridge IP Label
        label_bridge_ip = tk.Label(self.root, text="Hue Bridge IP Address")  

        self.hue_bridge_ip_text_entry = tk.StringVar()
        entry_bridge_ip = tk.Entry(self.root, textvariable=self.hue_bridge_ip_text_entry)

        label_bridge_ip.grid(row=0, column=0)
        entry_bridge_ip.grid(row=0, column=1)

        self.hue_bridge_ip_text_entry.set(hue_bridge_ip)


    def processUserChange(self, hue_bridge_ip, entry_light_name):
        self.hbl4z.save_setting("hue_bridge_ip", hue_bridge_ip)
        self.hbl4z.save_setting("hue_light_id", entry_light_name[:-1][entry_light_name.rfind("(id ") + 4:])
 

    def connect_to_bridge(self):
       
        # if we don't have the bridge IP in registry, try to find from Hue discovery
        if self.hbl4z.app_settings['hue_bridge_ip'] == None:
            print("Performing initial Bridge Discovery, please disconnect from any VPN.")
            # Attempt to find Hue Bridge IP using Philips Hue Discovery method

            try:
                response = requests.get("https://discovery.meethue.com")
                hue_discovery_result = response.json()

                if len(hue_discovery_result) > 0:
                    self.hbl4z.app_settings['hue_bridge_ip'] = hue_discovery_result[0]['internalipaddress']
                    print(f"Found Hue Bridge at IP {self.hbl4z.app_settings['hue_bridge_ip']}...")

                # Save Bridge IP
                self.hbl4z.save_setting("hue_bridge_ip", self.hbl4z.app_settings['hue_bridge_ip'])
            except:
                print("Unable to automatically discover Hue Bridge IP. Please enter manually...")

        self.draw_bridge_ip(self.hbl4z.app_settings['hue_bridge_ip'])

        # Authenticate to bridge
        if self.hbl4z.app_settings['hue_bridge_ip'] != None and self.hbl4z.app_settings['hue_username'] == None:
            # we don't have a username in registry, so prompt user to enroll using link button
            try:
                response = requests.post(f"http://{self.hbl4z.app_settings['hue_bridge_ip']}/api",json={"devicetype":"hue_busy_lamp_zoom"})
                api_response = response.json()
                self.hbl4z.app_settings['hue_username'] = api_response[0]['success']['username']
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
                self.hbl4z.save_setting("hue_username", self.hbl4z.app_settings['hue_username'])
        
        # Light ID

        if self.hbl4z.app_settings['hue_bridge_ip'] != None and self.hbl4z.app_settings['hue_username'] != None:
            print(f"Searching for lights using Hue Bridge API....")

            try:
                light_get_response = requests.get(url = f"http://{self.hbl4z.app_settings['hue_bridge_ip']}/api/{self.hbl4z.app_settings['hue_username']}/lights")
                light_response_json = light_get_response.json()
            except:
                print(f"Unable to Connect to Bridge...")
            else:
                
                hue_light_dict = {}

                if len(light_response_json) > 0:
                    for light in light_response_json:
                        if "colormode" in light_response_json[light]["state"]:
                            hue_light_dict[light] = light_response_json[light]['name']
                
                hue_light_list = [f"{light_response_json[light]['name']} (id {light})" for light in light_response_json if "colormode" in light_response_json[light]["state"]]
                
                if light_get_response.status_code==200:
                    print(f"Successfully connected to bridge, found {len(hue_light_list)} color capable bulbs.")
                    self.connected_to_bridge = True
                else:
                    print("Error: Unable to retrieve list of bulbs from bridge...")

                if self.hbl4z.app_settings['hue_light_id'] == "" or self.hbl4z.app_settings['hue_light_id'] == None:
                    hue_selected_light_name = None
                    print("Please select a light above and click Save.")
                else:
                    hue_selected_light_name = f"{hue_light_dict[self.hbl4z.app_settings['hue_light_id']]} (id {self.hbl4z.app_settings['hue_light_id']})"
                
                self.draw_hue_light_list(hue_light_list,hue_selected_light_name)

 

    def tk_mainloop(self):
        self.root.title("Hue Busy Light for Zoom")

        # Save button
        button = tk.Button(
            text="Save Changes",
            command=lambda: self.processUserChange(self.hue_bridge_ip_text_entry.get(), self.entry_light_name_text_entry.get()),
        )

        # Display console logging data inside tkinter
        console_logging_window = scrolledtext.ScrolledText(self.root, undo=True)
        console_output_window = PrintLogger(console_logging_window)
        sys.stdout = console_output_window

        # Layout grid
        button.grid(row=2, columnspan=2)
        console_logging_window.grid(row=3, columnspan=2)

        # Start Workers

        t1 = threading.Thread(target=self.run_t1)
        t1.start()

        t2 = threading.Thread(target=self.run_t2)
        t2.start()

        self.root.mainloop()
        
        self.stop_workers = True
        
        t1.join()
        t2.join()
        
    def run_t1(self):        
        while not self.stop_workers:
            self.hbl4z.zoom_status_monitor()
            time.sleep(2)

    def run_t2(self):
        # long process goes here
        while not self.connected_to_bridge and not self.stop_workers:
            self.connect_to_bridge()
            time.sleep(5)


if __name__ == "__main__":
    app = tk_hue_light_for_zoom_app()
    
    