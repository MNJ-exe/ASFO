import os
import shutil
import sys
import subprocess
import winreg as reg

# Function to organize files in the specified folder based on file extensions
def clean_folder(folder_path):
    if os.path.isdir(folder_path):
        # Iterate through all files in the folder
        for file_name in os.listdir(folder_path):
            full_file_path = os.path.join(folder_path, file_name)
            if os.path.isfile(full_file_path):
                # Get the file extension and create a subfolder name
                file_extension = file_name.split('.')[-1].lower()
                subfolder_name = f"{file_extension.upper()} Files"
                subfolder_path = create_subfolder_if_needed(folder_path, subfolder_name)
                new_location = os.path.join(subfolder_path, file_name)
                # Move the file if it does not already exist in the subfolder
                if not os.path.exists(new_location):
                    shutil.move(full_file_path, new_location)
                    print(f"Moved {file_name} to {subfolder_name}")
    else:
        print("Invalid folder path.")

# Function to create a subfolder if it does not already exist
def create_subfolder_if_needed(folder_path, subfolder_name):
    subfolder_path = os.path.join(folder_path, subfolder_name)
    if not os.path.exists(subfolder_path):
        os.mkdir(subfolder_path)
    return subfolder_path

# Function to add the application to the Windows right-click context menu
def add_to_registry():
    key_path = r"Directory\shell\ASFO"  # Main context menu key
    command_key_path = r"Directory\shell\ASFO\command"  # Command subkey
    
    # Create the main context menu item and set an icon
    with reg.CreateKey(reg.HKEY_CLASSES_ROOT, key_path) as key:
        reg.SetValue(key, "", reg.REG_SZ, "ASFO")  # Display name
        icon_path = r"C:\Users\Manoj\Pictures\asfoico.ico"  # Icon file path
        reg.SetValueEx(key, "Icon", 0, reg.REG_SZ, icon_path)
    
    # Create the command key to run this script with the selected folder
    with reg.CreateKey(reg.HKEY_CLASSES_ROOT, command_key_path) as command_key:
        executable_path = f'"{os.path.abspath(sys.argv[0])}" "%1"'
        reg.SetValue(command_key, "", reg.REG_SZ, executable_path)

# Function to uninstall the application by running an uninstall script
def uninstall():
    current_directory = os.path.dirname(os.path.abspath(sys.argv[0]))
    subprocess.run([sys.executable, os.path.join(current_directory, "uninstall.py")])

# Function to check if the application has been removed, triggering uninstall cleanup
def check_for_uninstall():
    current_executable = os.path.abspath(sys.argv[0])
    if not os.path.exists(current_executable):
        uninstall()  # Remove registry entry if executable no longer exists
        return True
    return False

# Main program execution
if __name__ == "__main__":
    if check_for_uninstall():
        sys.exit(0)  # Exit if the program was uninstalled

    # Handle command-line arguments
    if len(sys.argv) > 1:
        if sys.argv[1].lower() == "uninstall":
            uninstall()  # Run uninstall if 'uninstall' argument is passed
        else:
            folder_path = sys.argv[1]
            clean_folder(folder_path)  # Organize files in the specified folder
    else:
        add_to_registry()  # Add the application to right-click menu if no arguments
        print("Application installed and added to right-click menu.")
