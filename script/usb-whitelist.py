# Import:
import subprocess
import win32api
import win32con
import win32gui
import win32gui_struct
import threading
import time
import usb.backend.libusb1
import usb.core
import os
import json
import usb.util
import configparser

# Config
config = configparser.ConfigParser()
config.read('../config/settings.conf')
block = config['config']['block']
lock = config['config']['lock']
log = config['config']['log']

# USB Backend
current_dir = os.path.dirname(os.path.abspath(__file__))
backendpath = os.path.join(current_dir, "..", "backend", "libusb-1.0.dll")
backend = usb.backend.libusb1.get_backend(find_library=lambda x: backendpath)


# Constants
WM_DEVICECHANGE = 0x0219
DBT_DEVICEARRIVAL = 0x8000
DBT_DEVICEREMOVECOMPLETE = 0x8004
DBT_DEVNODES_CHANGED = 0x0007
DBT_DEVTYP_DEVICEINTERFACE = 0x0005
GUID_DEVINTERFACE_USB_DEVICE = "{A5DCBF10-6530-11D2-901F-00C04FB951ED}"
# TODO Randomize the name of this file:
file_path = "../usb_devices.json"



class USBDeviceMonitor:
    def __init__(self):
        self.hwnd = None

    def start_monitoring(self):
        message_map = {
            WM_DEVICECHANGE: self.on_device_change,
        }

        wc = win32gui.WNDCLASS()
        wc.lpfnWndProc = message_map
        wc.lpszClassName = "USBDeviceMonitor"
        wc.hInstance = win32api.GetModuleHandle(None)
        class_atom = win32gui.RegisterClass(wc)

        self.hwnd = win32gui.CreateWindow(class_atom, "USBDeviceMonitor", 0, 0, 0, 0, 0, 0, 0, wc.hInstance, None)
        self.register_device_notification(self.hwnd)

        while True:
            try:
                win32gui.PumpWaitingMessages()
            except Exception as e:
                print("An error occurred while pumping messages:", e)
            time.sleep(0.1)

    def log_info(self, message, device):
        if log == "Yes":
            log_file_path = "../log/log.txt"
            current_time = time.strftime("%Y-%m-%d %H:%M:%S")
            log_message = f"[{current_time}] {message}:\n {device}"
            with open(log_file_path, "a") as log_file:
                log_file.write(log_message)
        else:
            pass

    def lock_windows(self):
        if lock == "Yes":
            try:
                # Lock workstation
                subprocess.run(["rundll32.exe", "user32.dll,LockWorkStation"], check=True)
                self.log_info("session locked", ".")
            except subprocess.CalledProcessError as e:
                self.log.info("ERROR locking sessions: ", e)
        else:
            pass

    def disable_usb_device(self, vendor_id, product_id):
        # Convertir les IDs en format hexadécimal
        vendor_id = f"{vendor_id:04X}"
        product_id = f"{product_id:04X}"
        if block == "Yes":
            self.log_info("Disabled:", product_id)
            powershell_script = f"""
                $device = Get-PnpDevice | Where-Object {{ $_.InstanceId -like "*VID_{vendor_id}&PID_{product_id}*" }}
                if ($device) {{
                    Disable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false 
                    Write-Output "Device disabled: $($device.Name)"
                }} else {{
                    Write-Output "Device not found"
                }}
                """  
            # Call the powershell script:
            result = subprocess.run(["powershell", "-Command", powershell_script, ], capture_output=True, text=True, encoding='latin1', errors='ignore')
            print("stdout:", result.stdout)
        else:
            pass
          
    # Get usb devices list 
    def get_usb_device_list(self):
        devices = usb.core.find(find_all=True)
        usb_device_list = []

        if devices is None:
            print("Error: No USB devices detected.")
            return 

        for device in devices:
            usb_device_info = {
                "Vendor ID": device.idVendor,
                "Product ID": device.idProduct,
                "Serial Number": usb.util.get_string(device, device.iSerialNumber),
            }
            usb_device_list.append(usb_device_info)
        return usb_device_list
    
    # Load from whitelist
    def load_usb_devices_from_file(self, file_path):
        if os.path.exists(file_path):
            with open(file_path, "r") as file:
                return json.load(file)
        else:
            return []
    
    # Compare USB connected list to the whitelist:
    def compare_new_usb_device(self, file_path):
        existing_devices = self.load_usb_devices_from_file(file_path)
        new_devices = self.get_usb_device_list()

        # Compare
        for new_device in new_devices:
            if new_device not in existing_devices:
                #Intrusion:
                print("UNWHITELISTED USB DETECTED;")
                print(new_device)

                # Get ID 
                vendor_id = new_device['Vendor ID']
                product_id = new_device['Product ID']

                # Security actions:
                self.disable_usb_device(vendor_id, product_id)
                self.lock_windows()
                self.log_info("Intrusion detected: New USB device connected", new_device)
                    
            else:
                print("USB disconnected OR connected but whitelisted")


    def register_device_notification(self, hwnd):
        filter = win32gui_struct.PackDEV_BROADCAST_DEVICEINTERFACE(GUID_DEVINTERFACE_USB_DEVICE, "")
        win32gui.RegisterDeviceNotification(hwnd, filter, win32con.DEVICE_NOTIFY_WINDOW_HANDLE)

    def on_device_change(self, hwnd, msg, wparam, lparam):
        print(f"device change: wparam={wparam} hwnd={hwnd} msg={msg} lparam={lparam}")
        # Device connected OR deconnected Verify whitelist:
        self.compare_new_usb_device(file_path)
        if wparam == DBT_DEVICEARRIVAL:
            print("New device connected")
            self.handle_device_change(lparam)
        elif wparam == DBT_DEVICEREMOVECOMPLETE:
            print("Device removed")
            self.handle_device_change(lparam)
        elif wparam == DBT_DEVNODES_CHANGED:
            print("Device nodes changed")
        else:
            print(f"Unhandled device change event: wparam={wparam} hwnd={hwnd} msg={msg} lparam={lparam}")
        return True

if __name__ == "__main__":
    monitor = USBDeviceMonitor()
    monitor_thread = threading.Thread(target=monitor.start_monitoring)
    monitor_thread.start()
    while True:
        time.sleep(1)
     

        
