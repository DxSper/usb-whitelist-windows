import json
import usb.backend.libusb1
import usb.core
import sys
import usb.util
import os
import subprocess
import script.usb_recover_all as usb_recover_all
import ctypes 
import configparser
import win32com.client

#config
config_file = 'config/settings.conf'
config = configparser.ConfigParser()
config.read(config_file)

#backend
current_dir = os.path.dirname(os.path.abspath(__file__))
backendpath = os.path.join(current_dir, "backend", "libusb-1.0.dll")
backend = usb.backend.libusb1.get_backend(find_library=lambda x: backendpath)

usb_devices = usb.core.find(backend=backend, find_all=True)

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def has_usb_devices(file_path):
    # Vérifie si le fichier JSON contient des éléments
    if os.path.exists(file_path):
        with open(file_path, 'r') as file:
            data = json.load(file)
            return len(data) > 0
    else:
        return False

def exec_monitor():
    print("Monitor is running")
    whitelist_script_path = os.path.join(current_dir, "script", "usb-whitelist.py")
    subprocess.run(["python", whitelist_script_path], shell=False)

#start monitor
def start_monitor():
    if is_admin() == True:
        if has_usb_devices: 
            exec_monitor()
        else:
            warning_nowhitelist = input("No whitelist configured. Maybe everythings gonna be blocked. \n Do you want to Start Monitor? (y/n)")
            if warning_nowhitelist == "y":
                exec_monitor()
            else: 
                print("Execution canceled.")
    else:
        print("You didn't run the script as Administrator. So the USB blocking functionality won't work, but the locking and logging features are still operational.")
        warning_noadmin = input("Continue ? (y/n)")
        if warning_noadmin == "y":
            exec_monitor()
        else:
            print("Execution canceled")

# Load whitelist
def load_usb_devices_from_file(file_path):
    if os.path.exists(file_path):
        with open(file_path, "r") as file:
            return json.load(file)
    else:
        return []

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

# Add device to whitelist 
def add_new_usb_devices(file_path):
    isEmpty = False

    try:
        existing_devices = load_usb_devices_from_file(file_path)
    except:
        print("file is corrupted or empty")
        datareset = []
        factoreset = input("Do you want to reset the whitelist ? (yes/no)")
        if factoreset == "yes":
            with open(file_path, "w") as file:
                json.dump(datareset, file)
            print("Reset of the whitelist done.")
            isEmpty = True
        elif factoreset == "no":
            print(f"{file_path} is empty or corrupted. Please check it. ERROR:")
        else:
            print(f"{file_path} is empty or corrupted. Please check it. ERROR:")

    new_devices = get_usb_device_list()

    if isEmpty == True:
        existing_devices = load_usb_devices_from_file(file_path)
    
    # Remove duplicates
    for new_device in new_devices:
        if new_device not in existing_devices:
            
            print(f"{new_device} added to the whitelist")
            existing_devices.append(new_device)
    with open(file_path, "w") as file:
        json.dump(existing_devices, file, indent=4)
             
# Reset whitelist and get a new whitelist
def reset_usb_devices_to_file(file_path):
    clear_json_file(file_path)
    usb_devices = list(get_usb_device_list())
    with open(file_path, "w") as file:
        json.dump(usb_devices, file, indent=4)
        print("Whitelisting new devices [OK]")

# Delete whitelist and dont trust any usb
def clear_json_file(file_path):
    with open(file_path, "w") as file:
        json.dump([], file)

# Config manager  
def config_manager():
    os.system("cls")
    print("Config manager:")

    block_q = "Do you want to block unwhitelisted USB device ? default=Yes"
    config['config']['block'] = configchoice_input(block_q)

    lock_q = "Do you want to lock the computer when unwhitelisted USB device is connected ? default=Yes" 
    config['config']['lock'] = configchoice_input(lock_q)

    log_q = "Do you want to log unwhitelisted USB device ? default=Yes" 
    config['config']['log'] = configchoice_input(log_q)

    with open(config_file, 'w') as configfile:
        config.write(configfile)
    
    print("Config succesfully saved at config/settings.conf")


def configchoice_input(q):
    configchoice = input(f"{q}\n(Yes/No) : ")
    if configchoice == "Yes":
        return "Yes"
    elif configchoice == "No":
        return "No"
    else:
        os.system("cls")
        return configchoice_input(q)

def choiceprint(options):
    print(f" {usbart}")
    
    for option in options:
        print(f"{option}")

options = [
    "1 - Whitelist all device connected             | 4 - Delete whitelist     ",
    "2 - Add a new device to whitelist              | 5 - Recover disabled device and whitelist",
    "3 - Reset whitelist and get a new whitelist    | 6 - Manage config  ",
    "7 - Start Monitor                              | e - exit      ",
] 

usbart = """      
                  888     
                  888      
888  888 .d8888b  88888b.  
888  888 88K      888 "88b 
888  888 "Y8888b. 888  888 
Y88b 888      X88 888 d88P 
 "Y88888  88888P' 88888P"  Whitelist
"""



def choice(file_path):
    choice = input("\nChoose [1] [2] [3] [4] [5] [6] [7] [e]: \n")

    if choice == "1":
        add_new_usb_devices(file_path)
    elif choice == "2":
        add_new_usb_devices(file_path)
    elif choice == "3":
        print("Whitelist [CLEARED]")
        reset_usb_devices_to_file(file_path)
    elif choice == "4":
        clear_json_file(file_path)
        print("Whitelist [CLEARED]")
    elif choice == "5":
        usb_recover_all.recover()
        print("Whitelist [OK]")
    elif choice == "6":
        config_manager()
    elif choice == "7":
        start_monitor()
    elif choice == "e":
        os.system("cls")
        sys.exit(0)
    else:
        os.system("cls")
        choiceprint(options)
        pass
        
def install_requirements(requirements_file):
    result = subprocess.run(['pip', 'install', '-r', requirements_file], capture_output=True, text=True)
    
    # Sucess
    if result.returncode == 0:
        print("All requirements are installed.")
    else:
        # error
        print("Error installing requirements:")
        print(result.stderr)

def add_task_admin():
    # Script path
    manager_path = os.path.abspath(__file__)
    script_directory = os.path.dirname(manager_path)
    monitor_script_path = os.path.join(script_directory, "usb-whitelist.py")
    # Lunch at starturp with admin right 
    subprocess.run(['schtasks', '/create', '/tn', 'MonitorStartup', '/tr', f'"{sys.executable}" "{monitor_script_path}"', '/sc', 'onstart', '/rl', 'highest', '/f'])

def add_task_user():
    #Script path
    manager_path = os.path.abspath(__file__)
    script_directory = os.path.dirname(manager_path)

    # Path to startup folder
    startup_folder = os.path.join(os.environ['APPDATA'], 'Microsoft\\Windows\\Start Menu\\Programs\\Startup')
    shortcut_path = os.path.join(startup_folder, 'hidden_monitor.lnk')
    bat_file_path = os.path.join(script_directory, "monitor_hidden.bat")

    # Create a shortcut 
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.Targetpath = bat_file_path
    shortcut.WorkingDirectory = script_directory
    shortcut.IconLocation = sys.executable
    shortcut.save()
    print("USB whitelisted NoAdmin is set to start as Startup")

# TODO remove start at startup admin 
# TODO remove start at 





if __name__ == "__main__":
    # Wait screen
    print("Starting USB Whitelist manager... The first start download requirements. Please wait.")
    # Install Or Verify requirements:
    install_requirements("requirements.txt")


    # TODO refractor in a fonciton
    # Vérifié si priv admin si non Demander les privilèges 
    if is_admin():
        print("Administrator privilege [OK]")

        # TODO: verify if its already starting at startup
        # Start at Startup
        startatstartup = input("Do you want to start the program at Startup with admin right ? (y/n)")
        if startatstartup == "y":
            add_task_admin()
        else:
            pass
    else:
        print("Please run as administrator to have full functionnality.")
        runasadmin = input("Do you want to run as administrator ? (y/n)")
        if runasadmin == "y":
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{__file__}"', None, 1)
        else:
            # TODO verify if its already starting at staturp
            startatstartupnoadmin = input("Do you want to start the program at Startup with no right (y/n)")
            if startatstartupnoadmin == "y":
                add_task_user()
            else:
                pass
    

    # TODO: randomize the name of this file:
    file_path = "usb_devices.json"
    choiceprint(options)
    while True:
        choice(file_path)

