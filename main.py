import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import subprocess
import threading
import os
from pathlib import Path

class SoftwareInstaller:
    def __init__(self, root):
        self.root = root
        self.root.title("Windows Software Installer")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        self.setup_ui()
        self.load_available_software()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="Windows Software Installer", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Available software list
        ttk.Label(main_frame, text="Available Software:").grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        
        self.software_listbox = tk.Listbox(main_frame, selectmode=tk.MULTIPLE, height=10)
        self.software_listbox.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        
        # Scrollbar for listbox
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.software_listbox.yview)
        scrollbar.grid(row=2, column=1, sticky=(tk.N, tk.S))
        self.software_listbox.config(yscrollcommand=scrollbar.set)
        
        # Selected software frame
        selected_frame = ttk.LabelFrame(main_frame, text="Selected for Installation", padding="5")
        selected_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(10, 5))
        
        self.selected_listbox = tk.Listbox(selected_frame, height=4)
        self.selected_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # Buttons frame
        buttons_frame = ttk.Frame(selected_frame)
        buttons_frame.grid(row=1, column=0, pady=(5, 0))
        
        ttk.Button(buttons_frame, text="Add Selected", command=self.add_selected).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(buttons_frame, text="Remove Selected", command=self.remove_selected).grid(row=0, column=1, padx=5)
        ttk.Button(buttons_frame, text="Clear All", command=self.clear_selected).grid(row=0, column=2, padx=(5, 0))
        
        # Installation options
        options_frame = ttk.LabelFrame(main_frame, text="Installation Options", padding="5")
        options_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(10, 5))
        
        ttk.Checkbutton(options_frame, text="Silent Installation (no prompts)").grid(row=0, column=0, sticky=tk.W)
        ttk.Checkbutton(options_frame, text="Install for all users").grid(row=1, column=0, sticky=tk.W)
        
        # Progress and output
        output_frame = ttk.LabelFrame(main_frame, text="Installation Progress", padding="5")
        output_frame.grid(row=5, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        self.output_text = scrolledtext.ScrolledText(output_frame, height=12, state=tk.DISABLED)
        self.output_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Install button
        self.install_button = ttk.Button(main_frame, text="Start Installation", 
                                       command=self.start_installation, state=tk.NORMAL)
        self.install_button.grid(row=6, column=0, pady=(15, 10))
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=7, column=0, sticky=(tk.W, tk.E))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        main_frame.rowconfigure(5, weight=1)
        output_frame.columnconfigure(0, weight=1)
        output_frame.rowconfigure(0, weight=1)
        
    def load_available_software(self):
        """Load predefined software options"""
        software_options = [
            "Google Chrome",
            "Mozilla Firefox",
            "VLC Media Player",
            "7-Zip",
            "Notepad++",
            "Python",
            "Visual Studio Code",
            "Git",
            "Node.js",
            "Adobe Reader",
            "WinRAR",
            "CCleaner",
            "Spotify",
            "Discord",
            "Zoom"
        ]
        
        for software in software_options:
            self.software_listbox.insert(tk.END, software)
    
    def add_selected(self):
        """Add selected software to installation list"""
        selected_indices = self.software_listbox.curselection()
        for index in selected_indices:
            software = self.software_listbox.get(index)
            if software not in self.selected_listbox.get(0, tk.END):
                self.selected_listbox.insert(tk.END, software)
    
    def remove_selected(self):
        """Remove selected software from installation list"""
        selected_indices = self.selected_listbox.curselection()
        for index in reversed(selected_indices):
            self.selected_listbox.delete(index)
    
    def clear_selected(self):
        """Clear all selected software"""
        self.selected_listbox.delete(0, tk.END)
    
    def get_install_command(self, software_name):
        """Return the appropriate installation command for each software"""
        commands = {
            "Google Chrome": 'winget install Google.Chrome --silent',
            "Mozilla Firefox": 'winget install Mozilla.Firefox --silent',
            "VLC Media Player": 'winget install VideoLAN.VLC --silent',
            "7-Zip": 'winget install 7zip.7zip --silent',
            "Notepad++": 'winget install Notepad++.Notepad++ --silent',
            "Python": 'winget install Python.Python.3 --silent',
            "Visual Studio Code": 'winget install Microsoft.VisualStudioCode --silent',
            "Git": 'winget install Git.Git --silent',
            "Node.js": 'winget install OpenJS.NodeJS --silent',
            "Adobe Reader": 'winget install Adobe.Acrobat.Reader.64-bit --silent',
            "WinRAR": 'winget install RARLab.WinRAR --silent',
            "CCleaner": 'winget install Piriform.CCleaner --silent',
            "Spotify": 'winget install Spotify.Spotify --silent',
            "Discord": 'winget install Discord.Discord --silent',
            "Zoom": 'winget install Zoom.Zoom --silent'
        }
        return commands.get(software_name, None)
    
    def log_output(self, message):
        """Add message to output text widget"""
        self.output_text.config(state=tk.NORMAL)
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.see(tk.END)
        self.output_text.config(state=tk.DISABLED)
    
    def update_status(self, message):
        """Update status bar"""
        self.status_var.set(message)
        self.root.update_idletasks()
    
    def install_software(self):
        """Install selected software"""
        selected_software = self.selected_listbox.get(0, tk.END)
        
        if not selected_software:
            messagebox.showwarning("Warning", "Please select at least one software to install.")
            return
        
        self.install_button.config(state=tk.DISABLED)
        self.log_output("Starting installation process...")
        self.update_status("Installing software...")
        
        success_count = 0
        fail_count = 0
        
        for software in selected_software:
            self.log_output(f"\nInstalling {software}...")
            self.update_status(f"Installing {software}...")
            
            command = self.get_install_command(software)
            if not command:
                self.log_output(f"Error: No installation command found for {software}")
                fail_count += 1
                continue
            
            try:
                # Run the installation command
                result = subprocess.run(command, shell=True, capture_output=True, text=True, timeout=300)
                
                if result.returncode == 0:
                    self.log_output(f"✓ {software} installed successfully!")
                    success_count += 1
                else:
                    self.log_output(f"✗ Failed to install {software}")
                    self.log_output(f"Error: {result.stderr}")
                    fail_count += 1
                    
            except subprocess.TimeoutExpired:
                self.log_output(f"✗ Timeout while installing {software}")
                fail_count += 1
            except Exception as e:
                self.log_output(f"✗ Error installing {software}: {str(e)}")
                fail_count += 1
        
        # Final summary
        self.log_output(f"\n=== Installation Complete ===")
        self.log_output(f"Successfully installed: {success_count}")
        self.log_output(f"Failed: {fail_count}")
        
        if fail_count == 0:
            self.update_status("All installations completed successfully!")
            messagebox.showinfo("Success", "All software installed successfully!")
        else:
            self.update_status(f"Installation completed with {fail_count} failure(s)")
            messagebox.showwarning("Completed", 
                                 f"Installation completed with {fail_count} failure(s). "
                                 f"Check the log for details.")
        
        self.install_button.config(state=tk.NORMAL)
    
    def start_installation(self):
        """Start installation in a separate thread to keep GUI responsive"""
        if not self.selected_listbox.get(0, tk.END):
            messagebox.showwarning("Warning", "Please select software to install first.")
            return
        
        # Confirm before starting
        if messagebox.askyesno("Confirm Installation", 
                             "This will install the selected software. Continue?"):
            # Run installation in separate thread
            thread = threading.Thread(target=self.install_software, daemon=True)
            thread.start()

def check_winget_available():
    """Check if winget (Windows Package Manager) is available"""
    try:
        result = subprocess.run("winget --version", shell=True, capture_output=True, text=True)
        return result.returncode == 0
    except:
        return False

def main():
    # Check if winget is available
    if not check_winget_available():
        messagebox.showerror("Error", 
                           "Windows Package Manager (winget) is not available.\n\n"
                           "Please install winget from the Microsoft Store or "
                           "ensure you're running Windows 10 1709+ or Windows 11.")
        return
    
    root = tk.Tk()
    app = SoftwareInstaller(root)
    
    # Center the window
    root.eval('tk::PlaceWindow . center')
    
    root.mainloop()

if __name__ == "__main__":
    main()
