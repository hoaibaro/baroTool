import ctypes
import os
import subprocess
import psutil # type: ignore
import tkinter as tk
from tkinter import messagebox
import sys

# NOTE: Hàm kiểm tra quyền admin
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    # Nếu không có quyền admin, yêu cầu chạy lại với quyền admin
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    sys.exit()

# NOTE: [1] Hàm chạy tất cả các chức năng theo thứ tự
def run_all():
    # Hiển thị hộp thoại xác nhận
    if not messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn chạy tất cả các tính năng? Quá trình này có thể mất một thời gian."):
        return

    # Danh sách các hàm cần chạy theo thứ tự
    functions_to_run = [
        ("Kích hoạt Windows", activate_windows),
        ("Cài đặt múi giờ", power_options),
        ("Cài đặt phần mềm", install_software_commands),
        ("Kích hoạt Office", activate_office),
        ("Bật tính năng Windows", turn_on_features),
        ("Đổi tên thiết bị", rename_device),
    ]

    # Chạy từng hàm theo thứ tự
    for name, func in functions_to_run:
        try:
            messagebox.showinfo("Đang chạy", f"Đang thực hiện: {name}")
            func(show_dialog=False)  # Không hiển thị hộp thoại khi chạy tự động
        except Exception as e:
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra khi thực hiện {name}: {str(e)}")
            break

    messagebox.showinfo("Hoàn thành", "Đã hoàn thành tất cả các tính năng!")

# NOTE: Hàm sao chép và cài đặt phần mềm
def copy_install():
    """Copy and install all required software"""
    # Tạo thư mục SETUP nếu chưa tồn tại
    if not os.path.exists(f"{os.environ['USERPROFILE']}\\Downloads\\SETUP"):
        os.makedirs(f"{os.environ['USERPROFILE']}\\Downloads\\SETUP")
        print("SETUP folder has been created successfully!")

    # Sao chép thư mục Software
    if not os.path.exists(f"{os.environ['USERPROFILE']}\\Downloads\\SETUP\\Software"):
        os.system("xcopy \"D:\\SOFTWARE\\PAYOO\\SETUP\" \"%USERPROFILE%\\Downloads\\SETUP\\Software\" /E /I /Y")
        print("Software folder has been copied successfully!")
    else:
        print("Software folder is already copied. Skipping...")

    # Sao chép thư mục Office2019
    if not os.path.exists(f"{os.environ['USERPROFILE']}\\Downloads\\SETUP\\Office2019"):
        os.system("xcopy \"D:\\SOFTWARE\\OFFICE\\Office 2019\\*\" \"%USERPROFILE%\\Downloads\\SETUP\\Office2019\" /E /I /Y")
        print("Office2019 folder has been copied successfully!")
    else:
        print("Office2019 folder is already copied. Skipping...")

    # Sao chép Unikey
    if not os.path.exists("C:\\unikey46RC2-230919-win64"):
        os.system("xcopy \"D:\\SOFTWARE\\PAYOO\\unikey46RC2-230919-win64\" \"C:\\unikey46RC2-230919-win64\" /E /H /C /I /Y")
        print("Unikey has been copied successfully!")
    else:
        print("Unikey is already copied. Skipping...")

    # Sao chép MSTeamsSetup
    if not os.path.exists("C:\\MSTeamsSetup.exe"):
        os.system("xcopy \"D:\\SOFTWARE\\PAYOO\\MSTeamsSetup.exe\" \"C:\\MSTeamsSetup.exe\" /E /H /C /I /Y")
        print("MSTeamsSetup has been copied successfully!")
    else:
        print("MSTeamsSetup is already copied. Skipping...")

    # Sao chép SC.exe
    if not os.path.exists(f"{os.environ['USERPROFILE']}\\Downloads\\SC-wKgXWicTb0XhUSNethaFN0vkhji53AY5mektJ7O_RSOdc8bEUVIEAAH_OewU.exe"):
        os.system("copy /Y \"D:\\SOFTWARE\\PAYOO\\SC-wKgXWicTb0XhUSNethaFN0vkhji53AY5mektJ7O_RSOdc8bEUVIEAAH_OewU.exe\" \"%USERPROFILE%\\Downloads\"")
        print("SC.exe has been copied successfully!")
    else:
        print("SC.exe is already copied. Skipping...")

    # Sao chép Trellix
    if not os.path.exists(f"{os.environ['USERPROFILE']}\\Downloads\\TrellixSmartInstall.exe"):
        os.system("copy /Y \"D:\\SOFTWARE\\PAYOO\\TrellixSmartInstall.exe\" \"%USERPROFILE%\\Downloads\"")
        print("Trellix has been copied successfully!")
    else:
        print("Trellix is already copied. Skipping...")

    # Sao chép MDMLaptop
    if not os.path.exists(f"{os.environ['USERPROFILE']}\\Downloads\\ManageEngine_MDMLaptopEnrollment"):
        os.system("xcopy /E /I /Y \"D:\\SOFTWARE\\PAYOO\\ManageEngine_MDMLaptopEnrollment\" \"%USERPROFILE%\\Downloads\\ManageEngine_MDMLaptopEnrollment\"")
        print("MDMLaptop has been copied successfully!")
    else:
        print("MDMLaptop is already copied. Skipping...")

    # Cài đặt 7-Zip
    if not os.path.exists(f"{os.environ['PROGRAMFILES']}\\7-Zip\\7z.exe"):
        os.system("start /wait \"%USERPROFILE%\\Downloads\\SETUP\\Software\\7z2408-x64.exe\" /S")
        print("7-Zip has been installed successfully!")
    else:
        print("7-Zip is already installed. Skipping...")

    # Cài đặt Google Chrome
    if not os.path.exists(f"{os.environ['PROGRAMFILES']}\\Google\\Chrome\\Application\\chrome.exe"):
        os.system("start /wait \"%USERPROFILE%\\Downloads\\SETUP\\Software\\ChromeSetup.exe\" /silent /install")
        print("Google Chrome has been installed successfully!")
    else:
        print("Google Chrome is already installed. Skipping...")

    # Cài đặt LAPS
    if not os.path.exists(f"{os.environ['PROGRAMFILES']}\\LAPS\\AdmPwd.UI.exe"):
        os.system("start /wait \"%USERPROFILE%\\Downloads\\SETUP\\Software\\LAPS_x64.msi\" /quiet")
        print("LAPS has been installed successfully!")
    else:
        print("LAPS is already installed. Skipping...")

    # Cài đặt Foxit PDF Reader
    if not os.path.exists(f"{os.environ['PROGRAMFILES']}\\Foxit Software\\Foxit PDF Reader\\FoxitReader.exe"):
        os.system("start /wait \"%USERPROFILE%\\Downloads\\SETUP\\Software\\FoxitPDFReader20243_enu_Setup_Prom.exe\" /silent /install")
        print("Foxit PDF Reader has been installed successfully!")
    else:
        print("Foxit PDF Reader is already installed. Skipping...")

    # Cài đặt Microsoft Office 2019
    if not os.path.exists(f"{os.environ['PROGRAMFILES']}\\Microsoft Office\\root\\Office16\\WINWORD.EXE"):
        print("Installing Microsoft Office 2019...")
        os.chdir(f"{os.environ['TEMP']}\\Office2019")
        os.system("setup.exe /configure configuration.xml")
        os.chdir(os.path.dirname(os.path.abspath(__file__)))
        print("MSOffice 2019 installed successfully!")
    else:
        print("MSOffice 2019 is already installed. Skipping...")

# NOTE: [2] Hàm cài đặt tất cả phần mềm cần thiết
def install_software_commands():
    """Install all required software"""
    # Tạo cửa sổ mới cho menu chọn thiết bị
    device_window = tk.Toplevel(root)
    device_window.title("Chọn thiết bị")
    device_window.geometry("400x300")
    device_window.configure(bg="black")

    # Tiêu đề menu
    title_label = tk.Label(device_window, text="SELECT DEVICE TYPE",
                           fg="green", bg="black", font=("Courier", 12, "bold"))
    title_label.pack(pady=10)

    # Tạo bảng chọn thiết bị
    table_frame = tk.Frame(device_window, bg="black")
    table_frame.pack(pady=20)

    # Tiêu đề cột
    headers = ["STT", "Thiết bị", "Mô tả"]
    for col, header in enumerate(headers):
        label = tk.Label(table_frame, text=header, fg="green", bg="black", 
                        font=("Arial", 10, "bold"), width=15)
        label.grid(row=0, column=col, padx=5, pady=5)

    # Dữ liệu bảng
    devices = [
        ["1", "Desktop", "Máy tính để bàn"],
        ["2", "Laptop", "Máy tính xách tay"]
    ]

    # Hiển thị dữ liệu
    for row, device in enumerate(devices, 1):
        for col, value in enumerate(device):
            label = tk.Label(table_frame, text=value, fg="white", bg="black", 
                           font=("Arial", 10), width=15)
            label.grid(row=row, column=col, padx=5, pady=5)

    # Frame cho nút chọn
    button_frame = tk.Frame(device_window, bg="black")
    button_frame.pack(pady=20)

    # Nút chọn Desktop
    desktop_btn = tk.Button(button_frame, text="[1] Cài đặt cho Desktop", 
                          command=lambda: [device_window.destroy(), install_for_desktop()],
                          font=("Arial", 10), fg="white", bg="green", width=20, height=2)
    desktop_btn.grid(row=0, column=0, padx=10)

    # Nút chọn Laptop
    laptop_btn = tk.Button(button_frame, text="[2] Cài đặt cho Laptop", 
                         command=lambda: [device_window.destroy(), install_for_laptop()],
                         font=("Arial", 10), fg="white", bg="green", width=20, height=2)
    laptop_btn.grid(row=0, column=1, padx=10)

    # Nút quay lại
    back_btn = tk.Button(button_frame, text="[0] Quay lại", 
                        command=device_window.destroy,
                        font=("Arial", 10), fg="white", bg="red", width=20, height=2)
    back_btn.grid(row=1, column=0, columnspan=2, pady=10)

# NOTE: Hàm cài đặt phần mềm cho máy Desktop
def install_for_desktop():
    """Install software for desktop"""
    messagebox.showinfo("Thông báo", "Đang cài đặt phần mềm cho Desktop...")
    copy_install()
    install_software_commands()
    messagebox.showinfo("Thông báo", "Cài đặt phần mềm cho Desktop hoàn tất!")

# NOTE: Hàm cài đặt phần mềm cho máy Laptop
def install_for_laptop():
    """Install software for laptop"""
    messagebox.showinfo("Thông báo", "Đang cài đặt phần mềm cho Laptop...")
    copy_install()
    install_software_commands()
    messagebox.showinfo("Thông báo", "Cài đặt phần mềm cho Laptop hoàn tất!")

# NOTE: [3] Hàm cấu hình Power Options và Firewall
def power_options():
    # Tạo cửa sổ mới cho menu
    def show_menu():
        menu_window = tk.Toplevel(root)
        menu_window.title("Power Options and Firewall")
        menu_window.geometry("400x300")
        menu_window.configure(bg="black")

        # Tiêu đề menu
        title_label = tk.Label(menu_window, text="SELECT DEVICE TYPE",
                               fg="green", bg="black", font=("Courier", 12, "bold"))
        title_label.pack(pady=10)

        # Các nút chức năng
        def set_time_and_power():
            # Setting Time and Timezone
            os.system("tzutil /s \"SE Asia Standard Time\"")
            os.system("reg add \"HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Services\\w32time\\Parameters\" /v Type /t REG_SZ /d NTP /f >nul")
            os.system("w32tm /resync >nul")
            os.system("reg add \"HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Services\\tzautoupdate\" /v Start /t REG_DWORD /d 2 /f >nul")
            messagebox.showinfo("Thông báo", "Thông tin cập nhật thành công!")
            # Turning off the firewall
            os.system("netsh advfirewall set allprofiles state off")
            messagebox.showinfo("Thông báo", "Firewall đã được tắt!")
            # Setting Power Options to "Do Nothing"
            os.system("powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_BUTTONS LIDACTION 0 >nul")
            os.system("powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_BUTTONS LIDACTION 0 >nul")
            os.system("powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_BUTTONS SBUTTONACTION 0 >nul")
            os.system("powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_BUTTONS SBUTTONACTION 0 >nul")
            os.system("powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_BUTTONS PBUTTONACTION 0 >nul")
            os.system("powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_BUTTONS PBUTTONACTION 0 >nul")
            os.system("powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_VIDEO VIDEOIDLE 0 >nul")
            os.system("powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_VIDEO VIDEOIDLE 0 >nul")
            os.system("powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_SLEEP STANDBYIDLE 0 >nul")
            os.system("powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_SLEEP STANDBYIDLE 0 >nul")
            os.system("powercfg /SETACTIVE SCHEME_CURRENT >nul")
            messagebox.showinfo("Thông báo", "Power Options setup completed successfully!")

        def turn_on_firewall():
            # Turning on the firewall
            os.system("netsh advfirewall set allprofiles state on")
            messagebox.showinfo("Thông báo", "Firewall đã được bật!")

        # Tạo các nút
        btn1 = tk.Button(menu_window, text="[1] Set Time/Timezone and Power Options", command=set_time_and_power,
                         font=("Arial", 10), fg="white", bg="green", width=40, height=2)
        btn1.pack(pady=5)

        btn2 = tk.Button(menu_window, text="[2] Turn on Firewall", command=turn_on_firewall,
                         font=("Arial", 10), fg="white", bg="green", width=40, height=2)
        btn2.pack(pady=5)

        btn3 = tk.Button(menu_window, text="[0] Return to Menu", command=menu_window.destroy,
                         font=("Arial", 10), fg="white", bg="red", width=40, height=2)
        btn3.pack(pady=5)

    # Hiển thị menu
    show_menu()

# NOTE: [4] Hàm quản lý ổ đĩa (đổi ký tự, thu nhỏ, gộp, đổi tên)
def change_drive_letter():
    def get_volume_name(drive_letter):
        kernel32 = ctypes.windll.kernel32
        volume_name_buffer = ctypes.create_unicode_buffer(1024)
        file_system_name_buffer = ctypes.create_unicode_buffer(1024)
        drive_path = f"{drive_letter}:\\"
        kernel32.GetVolumeInformationW(
            drive_path,
            volume_name_buffer,
            ctypes.sizeof(volume_name_buffer),
            None,
            None,
            None,
            file_system_name_buffer,
            ctypes.sizeof(file_system_name_buffer)
        )
        return volume_name_buffer.value or "No Label"

    def update_drives_list():
        drives_info = []
        for partition in psutil.disk_partitions():
            try:
                drive_letter = partition.device.strip(":\\")
                volume_name = get_volume_name(drive_letter)
                usage = psutil.disk_usage(partition.mountpoint)
                drives_info.append({
                    "drive": partition.device,
                    "volume_name": volume_name,
                    "size": f"{usage.total // (1024**3)} GB",
                    "free": f"{usage.free // (1024**3)} GB"
                })
            except Exception as e:
                print(f"Error reading drive {partition.device}: {str(e)}")
        return drives_info

    def show_drives_list():
        drives_info = update_drives_list()
        drives_listbox.delete(0, tk.END)
        for drive in drives_info:
            drives_listbox.insert(tk.END, f"{drive['drive']} - Volume: {drive['volume_name']} - Size: {drive['size']} - Free: {drive['free']}")

    def change_letter():
        # Xóa các widget cũ
        for widget in content_frame.winfo_children():
            widget.destroy()
        
        # Hiển thị danh sách ổ đĩa
        show_drives_list()
        
        # Tạo frame cho nhập ký tự mới
        input_frame = tk.Frame(content_frame, bg="black")
        input_frame.pack(pady=10)
        
        new_letter_var = tk.StringVar()
        new_letter_label = tk.Label(input_frame, text="Ký tự ổ đĩa mới (VD: Z):", fg="white", bg="black", font=("Arial", 10))
        new_letter_label.grid(row=0, column=0, padx=5)
        new_letter_entry = tk.Entry(input_frame, textvariable=new_letter_var, font=("Arial", 12), width=10)
        new_letter_entry.grid(row=0, column=1, padx=5)
        
        def submit_change_letter():
            selected_drive = drives_listbox.get(tk.ACTIVE).split(" - ")[0]
            new_letter = new_letter_var.get().upper()
            if selected_drive and new_letter:
                with open("ChangeLetter.txt", "w") as script_file:
                    script_file.write(f"select volume {selected_drive.strip(':')}\n")
                    script_file.write(f"assign letter={new_letter}\n")
                
                result = os.system("diskpart /s ChangeLetter.txt")
                os.remove("ChangeLetter.txt")
                
                if result == 0:
                    messagebox.showinfo("Thông báo", f"Đã đổi ký tự ổ đĩa từ {selected_drive} sang {new_letter} thành công!")
                    show_drives_list()
                else:
                    messagebox.showerror("Lỗi", "Không thể đổi ký tự ổ đĩa. Vui lòng kiểm tra lại.")
        
        submit_button = tk.Button(input_frame, text="Xác nhận", command=submit_change_letter, font=("Arial", 10), fg="white", bg="green", width=15)
        submit_button.grid(row=1, column=0, columnspan=2, pady=10)

    def shrink_volume():
        # Xóa các widget cũ
        for widget in content_frame.winfo_children():
            widget.destroy()
        
        # Hiển thị danh sách ổ đĩa
        show_drives_list()
        
        # Tạo frame cho các lựa chọn kích thước
        size_frame = tk.Frame(content_frame, bg="black")
        size_frame.pack(pady=10)
        
        size_label = tk.Label(size_frame, text="Chọn kích thước phân vùng mới:", fg="white", bg="black", font=("Arial", 10))
        size_label.pack()
        
        size_var = tk.StringVar()
        size_var.set("1")
        
        rb1 = tk.Radiobutton(size_frame, text="80GB (cho ổ 256GB)", variable=size_var, value="1", fg="green", bg="black", selectcolor="black")
        rb1.pack()
        rb2 = tk.Radiobutton(size_frame, text="200GB (cho ổ 500GB)", variable=size_var, value="2", fg="green", bg="black", selectcolor="black")
        rb2.pack()
        rb3 = tk.Radiobutton(size_frame, text="500GB (cho ổ 1TB+)", variable=size_var, value="3", fg="green", bg="black", selectcolor="black")
        rb3.pack()
        
        def submit_shrink():
            selected_drive = drives_listbox.get(tk.ACTIVE).split(" - ")[0]
            size_choice = size_var.get()
            
            size_mb = {
                "1": 82020,
                "2": 204955,
                "3": 512000
            }.get(size_choice)
            
            if selected_drive and size_mb:
                with open("ShrinkVolume.txt", "w") as script_file:
                    script_file.write(f"select volume {selected_drive.strip(':')}\n")
                    script_file.write(f"shrink desired={size_mb}\n")
                    script_file.write("create partition primary\n")
                    script_file.write("format fs=ntfs quick\n")
                    script_file.write("assign\n")
                
                result = os.system("diskpart /s ShrinkVolume.txt")
                os.remove("ShrinkVolume.txt")
                
                if result == 0:
                    messagebox.showinfo("Thông báo", "Đã thu nhỏ phân vùng thành công!")
                    show_drives_list()
                    rename_volume()
                else:
                    messagebox.showerror("Lỗi", "Không thể thu nhỏ phân vùng. Vui lòng kiểm tra lại.")
        
        submit_button = tk.Button(size_frame, text="Xác nhận", command=submit_shrink, font=("Arial", 10), fg="white", bg="green", width=15)
        submit_button.pack(pady=10)

    def merge_volume():
        # Xóa các widget cũ
        for widget in content_frame.winfo_children():
            widget.destroy()
        
        # Hiển thị danh sách ổ đĩa
        show_drives_list()
        
        # Tạo frame cho nhập thông tin
        input_frame = tk.Frame(content_frame, bg="black")
        input_frame.pack(pady=10)
        
        # Nhập ổ đĩa nguồn
        source_label = tk.Label(input_frame, text="Nhập ký tự ổ đĩa nguồn (VD: D):", fg="white", bg="black")
        source_label.grid(row=0, column=0, padx=5)
        source_entry = tk.Entry(input_frame, width=5)
        source_entry.grid(row=0, column=1, padx=5)
        
        # Nhập ổ đĩa đích
        target_label = tk.Label(input_frame, text="Nhập ký tự ổ đĩa đích (VD: C):", fg="white", bg="black")
        target_label.grid(row=1, column=0, padx=5, pady=5)
        target_entry = tk.Entry(input_frame, width=5)
        target_entry.grid(row=1, column=1, padx=5, pady=5)
        
        def submit_merge():
            source_drive = source_entry.get().upper()
            target_drive = target_entry.get().upper()
            
            if not source_drive or not target_drive:
                messagebox.showwarning("Cảnh báo", "Vui lòng nhập cả ổ đĩa nguồn và đích!")
                return
            
            ps_check = f"(Get-WmiObject Win32_Volume).DriveLetter -contains '{source_drive}:'"
            if subprocess.run(["powershell", "-command", ps_check], capture_output=True).returncode != 0:
                messagebox.showerror("Lỗi", f"Ổ đĩa nguồn {source_drive} không tồn tại!")
                return
            
            ps_check = f"(Get-WmiObject Win32_Volume).DriveLetter -contains '{target_drive}:'"
            if subprocess.run(["powershell", "-command", ps_check], capture_output=True).returncode != 0:
                messagebox.showerror("Lỗi", f"Ổ đĩa đích {target_drive} không tồn tại!")
                return
            
            with open("MergeVolume.txt", "w") as script_file:
                script_file.write(f"select volume {source_drive}\n")
                script_file.write("delete volume\n")
                script_file.write(f"select volume {target_drive}\n")
                script_file.write("extend\n")
            
            result = os.system("diskpart /s MergeVolume.txt")
            os.remove("MergeVolume.txt")
            
            if result == 0:
                messagebox.showinfo("Thông báo", "Đã gộp phân vùng thành công!")
                show_drives_list()
            else:
                messagebox.showerror("Lỗi", "Không thể gộp phân vùng. Vui lòng kiểm tra lại.")
        
        submit_button = tk.Button(input_frame, text="Xác nhận", command=submit_merge, font=("Arial", 10), fg="white", bg="green", width=15)
        submit_button.grid(row=2, column=0, columnspan=2, pady=10)

    def rename_volume():
        # Xóa các widget cũ
        for widget in content_frame.winfo_children():
            widget.destroy()
        
        # Hiển thị danh sách ổ đĩa
        show_drives_list()
        
        # Tạo frame cho nhập thông tin
        input_frame = tk.Frame(content_frame, bg="black")
        input_frame.pack(pady=10)
        
        # Nhập ổ đĩa cần đổi tên
        drive_label = tk.Label(input_frame, text="Nhập ký tự ổ đĩa cần đổi tên (VD: D):", fg="white", bg="black")
        drive_label.grid(row=0, column=0, padx=5)
        drive_entry = tk.Entry(input_frame, width=5)
        drive_entry.grid(row=0, column=1, padx=5)
        
        # Nhập tên mới
        name_label = tk.Label(input_frame, text="Nhập tên mới cho ổ đĩa:", fg="white", bg="black")
        name_label.grid(row=1, column=0, padx=5, pady=5)
        name_entry = tk.Entry(input_frame, width=20)
        name_entry.grid(row=1, column=1, padx=5, pady=5)
        
        def submit_rename():
            drive_letter = drive_entry.get().upper()
            new_label = name_entry.get()
            
            if not drive_letter:
                messagebox.showwarning("Cảnh báo", "Vui lòng nhập ký tự ổ đĩa!")
                return
            
            if not new_label:
                messagebox.showwarning("Cảnh báo", "Vui lòng nhập tên mới cho ổ đĩa!")
                return
            
            ps_check = f"(Get-WmiObject Win32_LogicalDisk).DeviceID -contains '{drive_letter}:'"
            if subprocess.run(["powershell", "-command", ps_check], capture_output=True).returncode != 0:
                messagebox.showerror("Lỗi", f"Ổ đĩa {drive_letter}: không tồn tại. Vui lòng kiểm tra lại.")
                return
            
            result = subprocess.run(["label", f"{drive_letter}:", new_label], capture_output=True, text=True)
            
            if result.returncode == 0:
                messagebox.showinfo("Thông báo", f"Đã đổi tên ổ đĩa {drive_letter}: thành \"{new_label}\"!")
                show_drives_list()
            else:
                messagebox.showerror("Lỗi", "Không thể đổi tên ổ đĩa. Vui lòng kiểm tra lại.")
        
        submit_button = tk.Button(input_frame, text="Xác nhận", command=submit_rename, font=("Arial", 10), fg="white", bg="green", width=15)
        submit_button.grid(row=2, column=0, columnspan=2, pady=10)

    # Tạo cửa sổ chính
    menu_window = tk.Toplevel(root)
    menu_window.title("Change Drive Letter / Edit Volume")
    menu_window.geometry("800x600")
    menu_window.configure(bg="black")

    # Tiêu đề menu
    title_label = tk.Label(menu_window, text="CHANGE THE DRIVE LETTER",
                           fg="green", bg="black", font=("Courier", 12, "bold"))
    title_label.pack(pady=10)

    # Tạo frame chứa danh sách ổ đĩa
    drives_frame = tk.Frame(menu_window, bg="black")
    drives_frame.pack(pady=10)

    drives_label = tk.Label(drives_frame, text="Danh sách ổ đĩa:", fg="white", bg="black", font=("Arial", 10))
    drives_label.pack()

    drives_listbox = tk.Listbox(drives_frame, font=("Courier", 10), width=70, height=10, bg="black", fg="green")
    drives_listbox.pack(pady=5)

    # Tạo frame chứa nội dung các chức năng
    content_frame = tk.Frame(menu_window, bg="black")
    content_frame.pack(pady=10)

    # Tạo các nút chức năng
    buttons_frame = tk.Frame(menu_window, bg="black")
    buttons_frame.pack(pady=10)

    btn1 = tk.Button(buttons_frame, text="[1] Change the drive letter", command=change_letter,
             font=("Arial", 10), fg="white", bg="green", width=20, height=2)
    btn1.grid(row=0, column=0, padx=5)

    btn2 = tk.Button(buttons_frame, text="[2] Shrink Volume", command=shrink_volume,
             font=("Arial", 10), fg="white", bg="green", width=20, height=2)
    btn2.grid(row=0, column=1, padx=5)

    btn3 = tk.Button(buttons_frame, text="[3] Merge Volume", command=merge_volume,
             font=("Arial", 10), fg="white", bg="green", width=20, height=2)
    btn3.grid(row=0, column=2, padx=5)

    btn4 = tk.Button(buttons_frame, text="[4] Rename Volume", command=rename_volume,
             font=("Arial", 10), fg="white", bg="green", width=20, height=2)
    btn4.grid(row=0, column=3, padx=5)

    btn5 = tk.Button(buttons_frame, text="[0] Return to Menu", command=menu_window.destroy,
             font=("Arial", 10), fg="white", bg="red", width=20, height=2)
    btn5.grid(row=0, column=4, padx=5)

    # Hiển thị danh sách ổ đĩa ban đầu
    show_drives_list()

# NOTE: [5] Hàm kích hoạt Windows
def activate_windows(show_dialog=True):
    if show_dialog:
        messagebox.showinfo("Thông báo", "Kích hoạt Windows và Office...")
    os.system("slmgr /ato")

# NOTE: [6] Hàm bật các tính năng Windows (.NET, IE)
def turn_on_features(show_dialog=True):
    if show_dialog:
        messagebox.showinfo("Thông báo", "Bật tính năng Windows (.NET, IE)...")
    os.system("dism /online /enable-feature /featurename:NetFx3")

# NOTE: [7] Hàm đổi tên thiết bị
def rename_device(show_dialog=True):
    if show_dialog:
        # Tạo cửa sổ mới
        rename_window = tk.Toplevel(root)
        rename_window.title("Đổi tên thiết bị")
        rename_window.geometry("400x300")
        rename_window.configure(bg="black")

        # Tiêu đề
        title_label = tk.Label(rename_window, text="Đổi tên thiết bị", fg="green", bg="black", font=("Arial", 12, "bold"))
        title_label.pack(pady=10)

        # Hiển thị username hiện tại
        current_name = os.environ.get('COMPUTERNAME', 'Unknown')
        current_name_label = tk.Label(rename_window, text=f"Tên thiết bị hiện tại: {current_name}", fg="white", bg="black", font=("Arial", 10))
        current_name_label.pack(pady=5)

        # Frame cho nhập tên mới
        input_frame = tk.Frame(rename_window, bg="black")
        input_frame.pack(pady=10)

        # Nhập tên mới
        new_name_label = tk.Label(input_frame, text="Nhập tên mới cho thiết bị:", fg="white", bg="black")
        new_name_label.grid(row=0, column=0, padx=5)
        new_name_entry = tk.Entry(input_frame, width=20)
        new_name_entry.grid(row=0, column=1, padx=5)

        def submit_rename():
            new_name = new_name_entry.get()
            if new_name:
                result = os.system(f'wmic computersystem where name="{current_name}" call rename name="{new_name}"')
                if result == 0:
                    messagebox.showinfo("Thông báo", "Tên thiết bị sẽ được đổi sau khi khởi động lại!")
                    rename_window.destroy()
                else:
                    messagebox.showerror("Lỗi", "Không thể đổi tên thiết bị. Vui lòng kiểm tra lại.")
            else:
                messagebox.showwarning("Cảnh báo", "Vui lòng nhập tên mới cho thiết bị!")

        # Nút xác nhận
        submit_button = tk.Button(rename_window, text="Xác nhận", command=submit_rename, font=("Arial", 10), fg="white", bg="green", width=15)
        submit_button.pack(pady=10)
    else:
        # Khi chạy tự động, sử dụng lệnh WMIC trực tiếp
        os.system('wmic computersystem where name="%COMPUTERNAME%" call rename name="%NewName%"')

# NOTE: [8] Hàm đặt mật khẩu cho user
def set_password():
    # Tạo cửa sổ mới để nhập mật khẩu
    password_window = tk.Toplevel(root)
    password_window.title("Đặt mật khẩu")
    password_window.geometry("400x200")
    password_window.configure(bg="black")

    # Tiêu đề
    title_label = tk.Label(password_window, text="Đặt mật khẩu mới", fg="green", bg="black", font=("Arial", 12, "bold"))
    title_label.pack(pady=10)

    # Hiển thị username hiện tại
    current_user = os.getlogin()
    username_label = tk.Label(password_window, text=f"Tên người dùng: {current_user}", fg="white", bg="black", font=("Arial", 10))
    username_label.pack(pady=5)

    # Entry để nhập mật khẩu mới
    new_password_var = tk.StringVar()
    new_password_entry = tk.Entry(password_window, textvariable=new_password_var, font=("Arial", 12), show="", width=25)
    new_password_entry.pack(pady=5)

    # Hàm xử lý đặt mật khẩu
    def submit_new_password():
        new_password = new_password_var.get()
        if new_password:
            result = os.system(f"net user {current_user} {new_password}")
            if result == 0:
                messagebox.showinfo("Thông báo", f"Mật khẩu đã được cập nhật thành công cho {current_user}!")
                password_window.destroy()
            else:
                messagebox.showerror("Lỗi", "Không thể cập nhật mật khẩu. Vui lòng kiểm tra quyền hạn của bạn.")
        else:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập mật khẩu mới!")

    # Nút xác nhận
    submit_button = tk.Button(password_window, text="Xác nhận", command=submit_new_password, font=("Arial", 10), fg="white", bg="green", width=15)
    submit_button.pack(pady=10)

# NOTE: [9] Hàm tham gia domain
def join_domain():
    os.system("SystemPropertiesComputerName")

# NOTE: [0] Hàm thoát chương trình và cài đặt múi giờ
def exit_program():
    root.quit()
    if show_dialog: # type: ignore
        messagebox.showinfo("Thông báo", "Đang cài đặt múi giờ Việt Nam...")
    os.system('tzutil /s "SE Asia Standard Time"')

# NOTE: Hàm kích hoạt Microsoft Office
def activate_office(show_dialog=True):
    if show_dialog:
        messagebox.showinfo("Thông báo", "Đang kích hoạt Office...")
    os.system('cscript "C:\\Program Files\\Microsoft Office\\Office16\\ospp.vbs" /inpkey:Q2NKY-J42YJ-X2KVK-9Q9PT-MKP63')
    os.system('cscript "C:\\Program Files\\Microsoft Office\\Office16\\ospp.vbs" /act')
    if show_dialog:
        messagebox.showinfo("Thông báo", "Office đã được kích hoạt!")

# Tạo cửa sổ chính
root = tk.Tk()
root.title("BAOPROVIP - Hệ thống quản lý")
root.geometry("700x400")  # Tăng chiều rộng để phù hợp với 2 cột
root.configure(bg="black")

# Tạo frame chứa tiêu đề
title_frame = tk.Frame(root, bg="black")
title_frame.grid(row=0, column=0, columnspan=2, pady=10, sticky="nsew")

# Tạo tiêu đề
title_label = tk.Label(title_frame, text="WELCOME TO BAOPROVIP", fg="green", bg="black", font=("Courier", 14, "bold"))
title_label.pack()

# Tạo các nút bấm cho từng tùy chọn
buttons = [
    # [1] Chạy tất cả các chức năng theo thứ tự
    ("Run All Options", run_all),
    # [2] Cài đặt tất cả phần mềm cần thiết
    ("Install All Software", install_software_commands),
    # [3] Cấu hình Power Options và Firewall
    ("Power Options and Firewall", power_options),
    # [4] Quản lý ổ đĩa (đổi ký tự, thu nhỏ, gộp, đổi tên)
    ("Change / Edit Volume", change_drive_letter),
    # [5] Kích hoạt Windows
    ("Activate", activate_windows),
    # [6] Bật các tính năng Windows (.NET, IE)
    ("Turn On Features", turn_on_features),
    # [7] Đổi tên thiết bị
    ("Rename Device", rename_device),
    # [8] Đặt mật khẩu cho user
    ("Set Password", set_password),
    # [9] Tham gia domain
    ("Join Domain", join_domain),
    # [0] Thoát chương trình và cài đặt múi giờ
    ("Exit", exit_program),
]

# Tính toán số hàng cho mỗi cột
rows_per_column = (len(buttons) + 1) // 2

# Thêm các nút vào giao diện theo grid layout
for index, (text, command) in enumerate(buttons):
    # Tính toán vị trí hàng và cột
    column = 0 if index < rows_per_column else 1
    row = index if index < rows_per_column else index - rows_per_column
    
    # Tạo và đặt vị trí nút
    btn = tk.Button(root, text=text, command=command, font=("Arial", 12), fg="white", bg="green", width=40, height=2)
    btn.grid(row=row+1, column=column, pady=5, padx=20, sticky="nsew")

# Cấu hình grid để căn giữa
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)
for i in range(rows_per_column + 1):
    root.grid_rowconfigure(i, weight=1)

# Chạy ứng dụng
root.mainloop()

