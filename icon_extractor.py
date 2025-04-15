import tkinter as tk
from tkinter import filedialog, ttk
from PIL import Image
import os
from pathlib import Path
import winreg
import win32gui
import win32con
import win32ui
import win32api
import logging
import webbrowser

class IconExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Icon Extractor")
        self.root.geometry("500x350")
        self.root.resizable(False, False)
        self.root.configure(bg="#e0e0e0")

        # Setup logging for debugging
        logging.basicConfig(filename='icon_extractor.log', level=logging.DEBUG,
                            format='%(asctime)s - %(levelname)s - %(message)s')

        # Default save path (User's Pictures/Icon Extractor)
        self.default_save_path = str(Path.home() / "Pictures" / "Icon Extractor")
        os.makedirs(self.default_save_path, exist_ok=True)
        self.save_path = self.default_save_path
        self.exe_path = None
        self.selected_format = tk.StringVar(value="PNG")
        self.selected_size = tk.StringVar(value="32x32")
        self.icon_sizes = []  # Store available icon sizes

        # Create UI
        self.create_widgets()

    def create_widgets(self):
        main_frame = tk.Frame(self.root, bg="#e0e0e0")
        main_frame.pack(pady=10, padx=20, fill="both", expand=True)

        tk.Label(main_frame, text="Icon Extractor", font=("Segoe UI", 16, "bold"), bg="#e0e0e0").pack(pady=10)

        self.file_name_label = tk.Label(main_frame, text="No file selected", font=("Segoe UI", 10), bg="#e0e0e0")
        self.file_name_label.pack(pady=5)

        # Frame for file selection buttons
        file_btn_frame = tk.Frame(main_frame, bg="#e0e0e0")
        file_btn_frame.pack(pady=5, fill="x")

        tk.Button(file_btn_frame, text="Select EXE File", font=("Segoe UI", 10), bg="#4CAF50", fg="white",
                command=self.select_exe_file, relief="flat", cursor="hand2").pack(side="left", expand=True, fill="x", padx=(0, 5))

        tk.Button(file_btn_frame, text="Select from Installed Apps", font=("Segoe UI", 10), bg="#2196F3", fg="white",
                command=self.select_installed_app, relief="flat", cursor="hand2").pack(side="left", expand=True, fill="x", padx=(5, 0))


        # Frame for Save Path
        path_frame = tk.Frame(main_frame, bg="#e0e0e0")
        path_frame.pack(pady=5, fill="x")

        tk.Label(path_frame, text="Save Icon To:", font=("Segoe UI", 10), bg="#e0e0e0").pack(side="left", padx=5)
        self.save_path_entry = tk.Entry(path_frame, font=("Segoe UI", 10))
        self.save_path_entry.insert(0, self.save_path)
        self.save_path_entry.pack(side="left", expand=True, fill="x", padx=5)

        tk.Button(main_frame, text="Change Save Path", font=("Segoe UI", 10), bg="#FFC107", fg="black",
                  command=self.change_save_path, relief="flat", cursor="hand2").pack(pady=5, fill="x")

        # Frame to hold Save Format and Icon Size side by side
        options_frame = tk.Frame(main_frame, bg="#e0e0e0")
        options_frame.pack(pady=5, fill="x")

        # Save Format
        tk.Label(options_frame, text="Save Format:", font=("Segoe UI", 10), bg="#e0e0e0").grid(row=0, column=0, padx=5, sticky="w")
        format_menu = ttk.Combobox(options_frame, textvariable=self.selected_format, values=["PNG", "ICO", "BMP"],
                                   state="readonly", font=("Segoe UI", 10), width=15)
        format_menu.grid(row=0, column=1, padx=5)

        # Icon Size
        tk.Label(options_frame, text="Icon Size:", font=("Segoe UI", 10), bg="#e0e0e0").grid(row=0, column=2, padx=5, sticky="w")
        self.size_menu = ttk.Combobox(options_frame, textvariable=self.selected_size, state="readonly", font=("Segoe UI", 10), width=15)
        self.size_menu.grid(row=0, column=3, padx=5)
        self.size_menu['values'] = ["32x32"]  # Default value
        self.size_menu.set("32x32")

        tk.Button(main_frame, text="Extract Icon", font=("Segoe UI", 10), bg="#FF5722", fg="white",
                  command=self.extract_icon, relief="flat", cursor="hand2").pack(pady=15, fill="x")
        
        # Info Button (GitHub link)
        info_btn = tk.Button(self.root, text="ℹ GitHub", font=("Segoe UI", 9), bg="#9E9E9E", fg="white",
                             command=self.open_github, relief="flat", cursor="hand2")
        info_btn.pack(side="bottom", pady=5)

    def open_github(self):
        webbrowser.open("https://github.com/MehdiSlr/Icon-Extractor")  # اینجا لینک گیت‌هابتو بزار

    def get_icon_sizes(self, exe_path):
        try:
            icon_count = win32gui.ExtractIconEx(exe_path, -1)
            if not icon_count or icon_count == 0:
                logging.warning(f"No icons found in {exe_path}")
                return []

            sizes = []
            large_icons, _ = win32gui.ExtractIconEx(exe_path, 0, icon_count)
            for hicon in large_icons:
                try:
                    icon_info = win32gui.GetIconInfo(hicon)
                    hbmColor = icon_info[4]
                    bmp_info = win32ui.CreateBitmapFromHandle(hbmColor).GetInfo()
                    width, height = bmp_info['bmWidth'], bmp_info['bmHeight']
                    sizes.append(f"{width}x{height}")
                except Exception as e:
                    logging.error(f"Error reading icon size: {str(e)}")
                finally:
                    win32gui.DestroyIcon(hicon)

            return sorted(set(sizes))  # Remove duplicates and sort
        except Exception as e:
            logging.error(f"Error getting icon sizes: {str(e)}")
            return []


    def select_exe_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Executable files", "*.exe")])
        if file_path:
            self.exe_path = file_path
            file_name = os.path.basename(file_path)
            self.file_name_label.config(text=f"Selected: {file_name}")
            logging.info(f"Selected EXE file: {file_path}")

            # Update icon sizes
            self.icon_sizes = self.get_icon_sizes(file_path)
            if not self.icon_sizes:
                self.icon_sizes = ["32x32"]
                self.file_name_label.config(text=f"Selected: {file_name} (No icons found, using default size)")
            self.size_menu['values'] = self.icon_sizes
            self.selected_size.set(self.icon_sizes[0] if self.icon_sizes else "32x32")

    def select_installed_app(self):
        app_window = tk.Toplevel(self.root)
        app_window.title("Installed Applications")
        app_window.geometry("400x300")
        app_window.configure(bg="#e0e0e0")

        app_listbox = tk.Listbox(app_window, font=("Segoe UI", 10), width=50, height=15)
        app_listbox.pack(pady=10, padx=10, fill="both", expand=True)

        apps = self.get_installed_apps()
        for app, path in apps.items():
            app_listbox.insert(tk.END, app)

        def on_select(event):
            selection = app_listbox.curselection()
            if selection:
                app_name = app_listbox.get(selection[0])
                self.exe_path = apps[app_name]
                self.file_name_label.config(text=f"Selected: {app_name}")
                logging.info(f"Selected installed app: {app_name}, Path: {self.exe_path}")

                # Update icon sizes
                self.icon_sizes = self.get_icon_sizes(self.exe_path)
                if not self.icon_sizes:
                    self.icon_sizes = ["32x32"]
                    self.file_name_label.config(text=f"Selected: {app_name} (No icons found, using default size)")
                self.size_menu['values'] = self.icon_sizes
                self.selected_size.set(self.icon_sizes[0] if self.icon_sizes else "32x32")

                app_window.destroy()

        app_listbox.bind("<<ListboxSelect>>", on_select)

    def clean_exe_path(self, raw_path):
        # مثال: '"C:\\Program Files\\App\\app.exe",0' → 'C:\\Program Files\\App\\app.exe'
        path = raw_path.strip().split(',')[0].strip('"').strip()
        return path if path.lower().endswith('.exe') else None

    def get_installed_apps(self):
        apps = {}
        registry_paths = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
            (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
        ]

        for hive, path in registry_paths:
            try:
                reg = winreg.ConnectRegistry(None, hive)
                key = winreg.OpenKey(reg, path)
                for i in range(1024):
                    try:
                        subkey_name = winreg.EnumKey(key, i)
                        subkey = winreg.OpenKey(key, subkey_name)
                        try:
                            app_name, _ = winreg.QueryValueEx(subkey, "DisplayName")
                            exe_path = None

                            # Try DisplayIcon first
                            try:
                                display_icon, _ = winreg.QueryValueEx(subkey, "DisplayIcon")
                                if display_icon:
                                    exe_path = self.clean_exe_path(display_icon)
                            except:
                                pass

                            # If not found, fallback to InstallLocation
                            if not exe_path:
                                try:
                                    install_location, _ = winreg.QueryValueEx(subkey, "InstallLocation")
                                    exe_path = self.find_exe_in_location(install_location, app_name)
                                except:
                                    pass

                            if app_name and exe_path and os.path.exists(exe_path):
                                apps[app_name] = exe_path
                        except:
                            pass
                        winreg.CloseKey(subkey)
                    except:
                        break
                winreg.CloseKey(key)
            except Exception as e:
                logging.error(f"Error accessing registry path {path}: {str(e)}")
        return apps


    def extract_exe_path_from_uninstall(self, uninstall_string):
        if not uninstall_string:
            return None
        uninstall_string = uninstall_string.strip('"')
        if os.path.exists(uninstall_string) and uninstall_string.lower().endswith('.exe'):
            return uninstall_string
        exe_dir = os.path.dirname(uninstall_string)
        for root, dirs, files in os.walk(exe_dir):
            for file in files:
                if file.lower().endswith('.exe'):
                    return os.path.join(root, file)
        return None

    def find_exe_in_location(self, install_location, app_name):
        if not install_location:
            return None
        for root, dirs, files in os.walk(install_location):
            for file in files:
                if file.lower().endswith('.exe'):
                    return os.path.join(root, file)
        return None

    def change_save_path(self):
        new_path = filedialog.askdirectory(initialdir=self.save_path)
        if new_path:
            self.save_path = new_path
            self.save_path_entry.delete(0, tk.END)
            self.save_path_entry.insert(0, self.save_path)
            logging.info(f"Save path changed to: {self.save_path}")

    def extract_icon(self):
        if not self.exe_path:
            self.file_name_label.config(text="Error: No file selected!")
            logging.error("No EXE file selected")
            return

        try:
            logging.info(f"Extracting icon from: {self.exe_path}")
            # Parse selected size
            size = self.selected_size.get()
            width, height = map(int, size.split('x'))

            # Get number of icons
            icon_count = win32gui.ExtractIconEx(self.exe_path, -1)
            if not icon_count or icon_count == 0:
                self.file_name_label.config(text="Error: No icon found in the EXE!")
                logging.error("No icon found in the EXE")
                return

            # Find the icon closest to the selected size
            hicon = None
            for i in range(icon_count):
                hicon_temp = win32gui.ExtractIconEx(self.exe_path, i)
                if hicon_temp and hicon_temp[0]:
                    hicon_temp = hicon_temp[0][0]
                    icon_info = win32gui.GetIconInfo(hicon_temp)
                    hbmColor = icon_info[4]
                    bmp_info = win32ui.CreateBitmapFromHandle(hbmColor).GetInfo()
                    if bmp_info['bmWidth'] == width and bmp_info['bmHeight'] == height:
                        hicon = hicon_temp
                        break
                    win32gui.DestroyIcon(hicon_temp)

            if not hicon:
                # If exact size not found, use the first icon
                hicon_temp = win32gui.ExtractIconEx(self.exe_path, 0)
                if hicon_temp and hicon_temp[0]:
                    hicon = hicon_temp[0][0]
                else:
                    self.file_name_label.config(text="Error: No icon found in the EXE!")
                    logging.error("No icon found in the EXE")
                    return

            # Convert icon to bitmap
            hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
            hbmp = win32ui.CreateBitmap()
            hbmp.CreateCompatibleBitmap(hdc, width, height)
            hdc = hdc.CreateCompatibleDC()
            hdc.SelectObject(hbmp)
            win32gui.DrawIconEx(hdc.GetHandleOutput(), 0, 0, hicon, width, height, 0, 0, win32con.DI_NORMAL)

            # Save bitmap to a BytesIO object
            bmp_info = hbmp.GetInfo()
            bmp_data = hbmp.GetBitmapBits(True)
            img = Image.frombytes('RGBA', (bmp_info['bmWidth'], bmp_info['bmHeight']), bmp_data, 'raw', 'BGRA')

            # Save icon in selected format
            file_name = os.path.splitext(os.path.basename(self.exe_path))[0]
            output_path = os.path.join(self.save_path, f"{file_name}.{self.selected_format.get().lower()}")

            if self.selected_format.get() == "PNG":
                img.save(output_path, "PNG")
            elif self.selected_format.get() == "ICO":
                img.save(output_path, "ICO")
            elif self.selected_format.get() == "BMP":
                img.save(output_path, "BMP")

            self.file_name_label.config(text=f"Icon saved to: {output_path}")
            logging.info(f"Icon saved to: {output_path}")

            # Clean up
            win32gui.DestroyIcon(hicon)

        except Exception as e:
            self.file_name_label.config(text=f"Error: {str(e)}")
            logging.error(f"Error extracting icon: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = IconExtractorApp(root)
    root.mainloop()