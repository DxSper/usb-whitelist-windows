# Import:
import subprocess
import time
import usb.backend.libusb1
import usb.core
import usb.util
import json
import os

# USB Backend
current_dir = os.path.dirname(os.path.abspath(__file__))
backendpath = os.path.join(current_dir, "..", "backend", "libusb-1.0.dll")
backend = usb.backend.libusb1.get_backend(find_library=lambda x: backendpath)

# TODO randomize name of this file:
file_path = "usb_devices.json"

# load whitelist
def load_usb_devices_from_file(file_path):
        if os.path.exists(file_path):
            with open(file_path, "r") as file:
                return json.load(file)
        else:
            return []

# Recover disabled devices:
def activate_all_usb_device(vendor_id, product_id):
    # Hex convert
    vendor_id = f"{vendor_id:04X}"
    product_id = f"{product_id:04X}"
    print(vendor_id, product_id)
    powershell_script = f"""
        $device = Get-PnpDevice | Where-Object {{ $_.InstanceId -like "*VID_{vendor_id}&PID_{product_id}*" }}
        if ($device) {{
            Enable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false 
            Write-Output "Device enabled: $($device.Name)"
        }} else {{
            Write-Output "Device not found"
        }}
        """
    # Call script PowerShell
    result = subprocess.run(["powershell", "-Command", powershell_script, ], capture_output=True, text=True, encoding='latin1', errors='ignore')
    print("stdout:", result.stdout)
    

# Get usb device list 
def get_usb_device_list():
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


def compare_new_usb_device():
    #Get devices
    devices = get_usb_device_list()
    #Get whitelisted devices
    existing_devices = load_usb_devices_from_file(file_path)
    
    # Detect no whitelisted device
    for device in devices:
        if device not in existing_devices:
            # Add to existing devices list
            existing_devices.append(device)
        else:
            print("Device already whitelisted")

    for existing_device in existing_devices:
        # Get ID 
        vendor_id = existing_device['Vendor ID']
        product_id = existing_device['Product ID']  
        activate_all_usb_device(vendor_id, product_id)        
    with open(file_path, "w") as file:
        json.dump(existing_devices, file, indent=4) 


def recover():
    print("All usb connected gonna be activated...")
    compare_new_usb_device()
    time.sleep(1)
    print("\nDone, all USB are enabled.")



# Do not execute this script here, use usb-white-list-manager to excute Recover All.



