import os
import sys
import winreg
import ctypes
from ctypes import windll
import subprocess

class ContextMenuHandler:
    def __init__(self):
        self.prefix_options = ["+Book+year+", "+Paper+year+", "+Thesis+year+", "+Report+year+", 
                               "+Slides+year+", "+Presentation+year+", "+Draft+year+"]
        self.suffix_options = ["+authors"]  
        self.menu_name = "Add Prefix-Suffix"
        self.python_executable = sys.executable
        self.script_path = os.path.abspath(__file__)
        self.pythonw_executable = self.python_executable.replace("python.exe", "pythonw.exe")
    
    def install(self):
        """Install the context menu entries in Windows Registry"""
        try:
            # Create main menu entry for files
            key_path = r"*\shell\{}".format(self.menu_name)
            key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, key_path)
            winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, self.menu_name)
            winreg.SetValueEx(key, "SubCommands", 0, winreg.REG_SZ, "")
            winreg.CloseKey(key)
            
            # Create submenu entries in command store
            shell_commands_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell"
            
            # Create prefix submenu entries
            for prefix_option in self.prefix_options:
                # Create each prefix option as a subcommand
                command_key = r"prefix.{}".format(prefix_option)
                command_key_path = os.path.join(shell_commands_path, command_key)
                
                key = winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, command_key_path)
                winreg.SetValueEx(key, None, 0, winreg.REG_SZ, "Add prefix '{}'".format(prefix_option))
                
                command_path = os.path.join(command_key_path, "command")
                command_key = winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, command_path)
                command_str = '"{}" "{}" prefix "{}" "%1"'.format(
                    self.pythonw_executable, self.script_path, prefix_option
                )
                winreg.SetValueEx(command_key, None, 0, winreg.REG_SZ, command_str)
                winreg.CloseKey(key)
                winreg.CloseKey(command_key)
            
            # Create suffix submenu entries
            for suffix_option in self.suffix_options:
                # Create each suffix option as a subcommand
                command_key = r"suffix.{}".format(suffix_option)
                command_key_path = os.path.join(shell_commands_path, command_key)
                
                key = winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, command_key_path)
                winreg.SetValueEx(key, None, 0, winreg.REG_SZ, "Add suffix '{}'".format(suffix_option))
                
                command_path = os.path.join(command_key_path, "command")
                command_key = winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, command_path)
                command_str = '"{}" "{}" suffix "{}" "%1"'.format(
                    self.pythonw_executable, self.script_path, suffix_option
                )
                winreg.SetValueEx(command_key, None, 0, winreg.REG_SZ, command_str)
                winreg.CloseKey(key)
                winreg.CloseKey(command_key)
            
            # Connect subcommands to main menu
            # We'll create a string of all command keys
            subcommands = []
            
            # Add prefix commands
            for prefix_option in self.prefix_options:
                subcommands.append(f"prefix.{prefix_option}")
            
            # Add suffix commands
            for suffix_option in self.suffix_options:
                subcommands.append(f"suffix.{suffix_option}")
            
            # Apply the subcommands to the main menu entry
            key_path = r"*\shell\{}".format(self.menu_name)
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path, 0, winreg.KEY_WRITE)
            winreg.SetValueEx(key, "SubCommands", 0, winreg.REG_SZ, ";".join(subcommands))
            winreg.CloseKey(key)
            
            print("Context menu entries installed successfully!")
            return True
            
        except Exception as e:
            print(f"Error installing context menu: {e}")
            return False
    
    def safe_delete_key(self, root_key, key_path):
        """Safely delete a registry key and all its subkeys"""
        try:
            # Try to open the key first to check if it exists
            try:
                winreg.OpenKey(root_key, key_path)
            except FileNotFoundError:
                return True  # Key doesn't exist, so no need to delete
            
            # First, delete all subkeys
            handle = winreg.OpenKey(root_key, key_path, 0, winreg.KEY_ALL_ACCESS)
            
            # Get subkey count
            try:
                info = winreg.QueryInfoKey(handle)
                num_subkeys = info[0]
                
                # Delete all subkeys
                for i in range(num_subkeys):
                    # Always get the first subkey since the index shifts after deletion
                    subkey_name = winreg.EnumKey(handle, 0)
                    subkey_path = f"{key_path}\\{subkey_name}"
                    self.safe_delete_key(root_key, subkey_path)
                
                # Close the handle to the key
                winreg.CloseKey(handle)
                
                # Delete the key itself now that all subkeys are gone
                winreg.DeleteKey(root_key, key_path)
                return True
                
            except Exception as e:
                winreg.CloseKey(handle)
                print(f"Failed to delete {key_path}: {e}")
                return False
                
        except Exception as e:
            print(f"Error accessing {key_path}: {e}")
            return False
    
    def uninstall(self):
        """Remove the context menu entries from Windows Registry"""
        try:
            successful = True
            
            # Remove main menu entry and its subkeys from HKEY_CLASSES_ROOT
            key_path = r"*\shell\{}".format(self.menu_name)
            if not self.safe_delete_key(winreg.HKEY_CLASSES_ROOT, key_path):
                successful = False
            
            # Remove submenu entries from HKEY_LOCAL_MACHINE
            shell_commands_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell"
            
            # Remove prefix submenu entries
            for prefix_option in self.prefix_options:
                command_key = r"prefix.{}".format(prefix_option)
                command_key_path = os.path.join(shell_commands_path, command_key)
                if not self.safe_delete_key(winreg.HKEY_LOCAL_MACHINE, command_key_path):
                    successful = False
            
            # Remove suffix submenu entries
            for suffix_option in self.suffix_options:
                command_key = r"suffix.{}".format(suffix_option)
                command_key_path = os.path.join(shell_commands_path, command_key)
                if not self.safe_delete_key(winreg.HKEY_LOCAL_MACHINE, command_key_path):
                    successful = False
            
            if successful:
                print("Context menu entries removed successfully!")
            else:
                print("Some registry keys could not be removed. You may need to delete them manually.")
            
            return successful
            
        except Exception as e:
            print(f"Error uninstalling context menu: {e}")
            return False
    
    def add_prefix(self, prefix, file_path):
        """Add a prefix to the selected file"""
        try:
            if not os.path.exists(file_path):
                return False
                
            # Get file directory, name and extension
            file_dir = os.path.dirname(file_path)
            file_name = os.path.basename(file_path)
            
            # Create new file path with prefix
            new_file_path = os.path.join(file_dir, f"{prefix}{file_name}")
            
            # Check if the new file name already exists
            counter = 1
            while os.path.exists(new_file_path):
                base_name, extension = os.path.splitext(file_name)
                new_file_path = os.path.join(file_dir, f"{prefix}{base_name} ({counter}){extension}")
                counter += 1
            
            # Rename the file with the prefix
            os.rename(file_path, new_file_path)
                        
            return True
            
        except Exception as e:
            print(f"Error adding prefix to file: {e}")
            return False
    
    def add_suffix(self, suffix, file_path):
        """Add a suffix to the selected file"""
        try:
            if not os.path.exists(file_path):
                return False
                
            # Get file directory, name and extension
            file_dir = os.path.dirname(file_path)
            file_name = os.path.basename(file_path)
            base_name, extension = os.path.splitext(file_name)
            
            # Create new file path with suffix
            new_file_path = os.path.join(file_dir, f"{base_name}{suffix}{extension}")
            
            # Check if the new file name already exists
            counter = 1
            while os.path.exists(new_file_path):
                new_file_path = os.path.join(file_dir, f"{base_name}{suffix} ({counter}){extension}")
                counter += 1
            
            # Rename the file with the suffix
            os.rename(file_path, new_file_path)
                        
            return True
            
        except Exception as e:
            print(f"Error adding suffix to file: {e}")
            return False


# Check if the script is running with admin privileges
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# Re-run the script with admin privileges if needed
def ensure_admin():
    if not is_admin():
        print("Elevating privileges...")
        windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, ' '.join([f'"{arg}"' for arg in sys.argv]), None, 1
        )
        sys.exit(0)

def main():
    handler = ContextMenuHandler()
    
    if len(sys.argv) > 1:
        command = sys.argv[1].lower()
        
        if command == "install":
            # Automatically elevate privileges for installation
            ensure_admin()
            handler.install()
            # input("Press Enter to exit...")
            
        elif command == "uninstall":
            # Automatically elevate privileges for uninstallation
            ensure_admin()
            handler.uninstall()
            # input("Press Enter to exit...")
            
        elif command == "prefix" and len(sys.argv) >= 4:
            prefix = sys.argv[2]
            file_path = sys.argv[3]
            handler.add_prefix(prefix, file_path)
        
        elif command == "suffix" and len(sys.argv) >= 4:
            suffix = sys.argv[2]
            file_path = sys.argv[3]
            handler.add_suffix(suffix, file_path)
                        
        else:
            print("Usage:")
            print("  - Install:         python script.py install")
            print("  - Uninstall:       python script.py uninstall")
            print("  - Add Prefix:      python script.py prefix [prefix_text] [file_path]")
            print("  - Add Suffix:      python script.py suffix [suffix_text] [file_path]")
            input("Press Enter to exit...")
    else:
        print("File Prefix/Suffix Context Menu")
        print("-------------------------------")
        print("Usage:")
        print("  - Install:         python script.py install")
        print("  - Uninstall:       python script.py uninstall")
        input("Press Enter to exit...")
        
if __name__ == "__main__":
    main()