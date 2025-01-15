import subprocess
import sys
import os
import ctypes

def is_admin():
    try:
        return os.getuid() == 0
    except AttributeError:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0

def run_powershell_script(script_path):
    # Set the execution policy scope for the current process to RemoteSigned and run the PowerShell script
    subprocess.run(["powershell", "-Command", "Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned;" + script_path], check=True)

if __name__ == "__main__":
    if is_admin():
        # Path to the existing PowerShell script
        ps_script_path = "365EmailListExport.ps1"

        # Run the PowerShell script
        run_powershell_script(ps_script_path)
    else:
        # Re-run the program with admin rights
        if sys.platform == 'win32':
            subprocess.run(['runas', '/user:Administrator', 'python'] + sys.argv)
        else:
            print("Please run the script as root.")
