import time
import ctypes
import tkinter as tk
from tkinter import ttk
import voicemeeterlib

class VoicemeeterController:
    def __init__(self, root):
        self.root = root
        self.root.title("VM Master Control")
        self.root.geometry("450x380")
        
        # --- UPDATE: Make window stay on top of others ---
        self.root.attributes('-topmost', True) 

        # 1. Connect to Voicemeeter
        try:
            self.vm = voicemeeterlib.api('banana')
            self.vm.login()
            print("Connected to Voicemeeter API")
            
            # Wait for connection to settle
            time.sleep(1)

            # Minimize Window
            self.force_minimize_window()

        except Exception as e:
            print(f"Connection Failed: {e}")

        # 2. Get FILTERED device lists (Only Real Hardware)
        self.clean_input_list = self.get_filtered_devices(is_input=True)
        self.clean_output_list = self.get_filtered_devices(is_input=False)

        # --- UI SECTION: CONFIGURATION ---
        self.lbl_setup = tk.Label(root, text="Step 1: Configuration", font=("Arial", 11, "bold"))
        self.lbl_setup.pack(pady=(15, 5))

        # A1 Output Selection
        tk.Label(root, text="Select Master Output (A1):", fg="#555").pack(pady=2)
        self.combo_a1 = ttk.Combobox(root, values=self.clean_output_list, width=45)
        self.combo_a1.pack(pady=2)
        self.combo_a1.set("Select Speakers...")

        # Mic Selection
        tk.Label(root, text="Select Microphone (Strip 0):", fg="#555").pack(pady=2)
        self.combo_mic = ttk.Combobox(root, values=self.clean_input_list, width=45)
        self.combo_mic.pack(pady=2)
        self.combo_mic.set("Select Mic...")

        # Master Apply Button
        self.btn_apply = tk.Button(root, text="APPLY & CONNECT DEVICES", 
                                   command=self.apply_settings, 
                                   bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
        self.btn_apply.pack(pady=15, ipadx=10, ipady=5)

        ttk.Separator(root, orient='horizontal').pack(fill='x', pady=10)

        # --- UI SECTION: LIVE CONTROL ---
        self.lbl_live = tk.Label(root, text="Step 2: Live Control (Strip 1 -> B1)", font=("Arial", 11, "bold"))
        self.lbl_live.pack(pady=(5, 5))

        # The Toggle Button
        self.btn_toggle = tk.Button(root, text="Loading...", 
                                    command=self.toggle_b1_routing, 
                                    width=20, font=("Arial", 10, "bold"))
        self.btn_toggle.pack(pady=5, ipady=5)

        # Status Bar
        self.status_label = tk.Label(root, text="System Ready", fg="grey")
        self.status_label.pack(side="bottom", pady=10)

        # Start the background checker
        self.update_ui_loop()

    def force_minimize_window(self):
        try:
            SW_MINIMIZE = 6
            user32 = ctypes.windll.user32
            hwnd = user32.FindWindowW(None, "Voicemeeter Banana")
            if hwnd:
                user32.ShowWindow(hwnd, SW_MINIMIZE)
        except Exception as e:
            print(f"Minimize Error: {e}")

    # --- NEW: FILTER FUNCTION ---
    def get_filtered_devices(self, is_input):
        """Returns a list of devices EXCLUDING Virtual/Cable devices"""
        devices = []
        
        # Determine if we are looking at Inputs or Outputs
        count = self.vm.device.ins if is_input else self.vm.device.outs
        getter = self.vm.device.input if is_input else self.vm.device.output

        for i in range(count):
            name = getter(i)['name']
            # FILTER LOGIC: Skip empty, Skip CABLE, Skip Voicemeeter
            if name and "CABLE" not in name and "Voicemeeter" not in name:
                devices.append(name)
        
        return devices

    def apply_settings(self):
        mic_name = self.combo_mic.get()
        a1_name = self.combo_a1.get()

        if "Select" in mic_name or "Select" in a1_name:
            self.status_label.config(text="Please select devices in the dropdowns first!", fg="red")
            return

        try:
            # 1. Set Hardware Devices
            self.vm.bus[0].device.wdm = a1_name
            self.vm.strip[0].device.wdm = mic_name
            
            # 2. Route Mic (Strip 0)
            self.vm.strip[0].A1 = False 
            self.vm.strip[0].B1 = True  
            
            # 3. Find and Connect VB-Cable (Internal Search)
            self.setup_cable_strip()

            self.status_label.config(text="Configuration Applied Successfully!", fg="green")
            self.update_toggle_button_visuals()

        except Exception as e:
            self.status_label.config(text=f"Error: {e}", fg="red")

    def setup_cable_strip(self):
        """
        Scans RAW device list to find the hidden 'CABLE Output'
        even though it is not in the dropdown menu.
        """
        cable_name = None
        
        # We loop through raw API inputs, not our filtered list
        for i in range(self.vm.device.ins):
            name = self.vm.device.input(i)['name']
            if "CABLE Output" in name:
                cable_name = name
                break
        
        if cable_name:
            self.vm.strip[1].device.wdm = cable_name
            self.vm.strip[1].A1 = True 
            self.vm.strip[1].B1 = True 
            print(f"Hidden Cable Found & Connected: {cable_name}")
        else:
            print("Error: Could not find 'CABLE Output' driver in Windows.")

    def toggle_b1_routing(self):
        try:
            if self.vm.pdirty: pass 
            current_state = self.vm.strip[1].B1
            new_state = not current_state
            self.vm.strip[1].B1 = new_state
            
            if new_state: 
                self.btn_toggle.config(text="DISCONNECT", bg="#ffcccc", fg="red")
            else: 
                self.btn_toggle.config(text="CONNECT", bg="#ccffcc", fg="green")
        except Exception as e:
            print(f"Toggle Error: {e}")

    def update_ui_loop(self):
        try:
            if self.vm.pdirty:
               self.update_toggle_button_visuals()
        except Exception:
            pass
        self.root.after(100, self.update_ui_loop)

    def update_toggle_button_visuals(self):
        is_connected = self.vm.strip[1].B1
        if is_connected:
            if self.btn_toggle['text'] != "DISCONNECT":
                self.btn_toggle.config(text="DISCONNECT", bg="#ffcccc", fg="red")
        else:
            if self.btn_toggle['text'] != "CONNECT":
                self.btn_toggle.config(text="CONNECT", bg="#ccffcc", fg="green")

    def cleanup(self):
        try:
            print("Shutting down Voicemeeter...")
            self.vm.command.shutdown()
            self.vm.logout()
        except Exception as e:
            print(f"Error during shutdown: {e}")
        finally:
            self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = VoicemeeterController(root)
    root.protocol("WM_DELETE_WINDOW", app.cleanup)
    root.mainloop()
