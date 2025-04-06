import os
import tkinter as tk
from tkinter import messagebox

# Hàm xử lý từng chức năng
def run_all():
    messagebox.showinfo("Thông báo", "Chạy tất cả các tùy chọn...")
    # Thêm các lệnh hệ thống cần chạy ở đây

def install_software():
    messagebox.showinfo("Thông báo", "Đang cài đặt phần mềm...")
    os.system("winget install Notepad++")

def power_options():
    messagebox.showinfo("Thông báo", "Cấu hình Power Options và Firewall...")
    os.system("powercfg /L")

def change_drive_letter():
    messagebox.showinfo("Thông báo", "Mở công cụ đổi ký tự ổ đĩa...")
    os.system("diskpart")

def activate_windows():
    messagebox.showinfo("Thông báo", "Kích hoạt Windows và Office...")
    os.system("slmgr /ato")

def turn_on_features():
    messagebox.showinfo("Thông báo", "Bật tính năng Windows (.NET, IE)...")
    os.system("dism /online /enable-feature /featurename:NetFx3")

def rename_device():
    new_name = tk.simpledialog.askstring("Đổi tên thiết bị", "Nhập tên mới:")
    if new_name:
        os.system(f"WMIC computersystem where caption='%COMPUTERNAME%' rename {new_name}")
        messagebox.showinfo("Thông báo", f"Thiết bị đã đổi thành {new_name}")

def set_password():
    username = tk.simpledialog.askstring("Đặt mật khẩu", "Nhập tên tài khoản:")
    password = tk.simpledialog.askstring("Đặt mật khẩu", "Nhập mật khẩu mới:", show="*")
    if username and password:
        os.system(f"net user {username} {password}")
        messagebox.showinfo("Thông báo", f"Mật khẩu đã cập nhật cho {username}")

def join_domain():
    domain = tk.simpledialog.askstring("Tham gia Domain", "Nhập tên domain:")
    if domain:
        os.system(f"netdom join {os.getenv('COMPUTERNAME')} /domain:{domain}")
        messagebox.showinfo("Thông báo", f"Đã tham gia domain {domain}")

def exit_program():
    root.quit()

# Tạo cửa sổ chính
root = tk.Tk()
root.title("BAOPROVIP - Hệ thống quản lý")
root.geometry("500x500")
root.configure(bg="black")

# Tạo tiêu đề
title_label = tk.Label(root, text="WELCOME TO BAOPROVIP", fg="green", bg="black", font=("Courier", 14, "bold"))
title_label.pack(pady=10)

# Tạo các nút bấm cho từng tùy chọn
buttons = [
    ("Run All Options", run_all),
    ("Install All Software", install_software),
    ("Power Options and Firewall", power_options),
    ("Change The Drive Letter / Edit Volume", change_drive_letter),
    ("Activate Windows 10 Pro / Office 2019 Pro Plus", activate_windows),
    ("Turn On Features (.NET Framework, Remove IE11)", turn_on_features),
    ("Rename Device", rename_device),
    ("Set Password", set_password),
    ("Join Domain", join_domain),
    ("Exit", exit_program),
]

# Thêm các nút vào giao diện
for text, command in buttons:
    btn = tk.Button(root, text=text, command=command, font=("Arial", 12), fg="white", bg="green", width=40, height=2)
    btn.pack(pady=5)

# Chạy ứng dụng
root.mainloop()

