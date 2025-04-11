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
    # Unset PSModulePath environment variable
    env = os.environ.copy()
    env.pop('PSModulePath', None)
    
    # Get the absolute path of the PowerShell script
    abs_script_path = os.path.abspath(script_path)
    
    # Import the Microsoft.PowerShell.Security module and set the execution policy
    command = f"Import-Module Microsoft.PowerShell.Security; Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned; & '{abs_script_path}'"
    
    # Run the PowerShell command
    subprocess.run(["powershell", "-Command", command], check=True, env=env)

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
