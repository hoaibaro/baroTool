## Hi there 👋

<!--
**hoaibaro/hoaibaro** is a ✨ _special_ ✨ repository because its `README.md` (this file) appears on your GitHub profile.

Here are some ideas to get you started:

- 🔭 I'm currently working on ...
- 🌱 I'm currently learning ...
- 👯 I'm looking to collaborate on ...
- 🤔 I'm looking for help with ...
- 💬 Ask me about ...
- 📫 How to reach me: ...
- 😄 Pronouns: ...
- ⚡ Fun fact: ...
-->

# BAOPROVIP - Công cụ quản lý hệ thống

## Giới thiệu
Đây là công cụ quản lý hệ thống cho Windows, cung cấp các tính năng:
- Cài đặt phần mềm (sử dụng winget)
- Quản lý nguồn điện và Firewall
- Quản lý ổ đĩa và phân vùng
- Kích hoạt Windows và Office
- Bật các tính năng Windows
- Đổi tên thiết bị
- Đặt mật khẩu
- Tham gia Domain
- Và nhiều tính năng khác

## Cách sử dụng

### Đối với người dùng cuối
1. Tải file `install.exe` từ thư mục `dist`
2. Chạy file `install.exe`
   - Chương trình sẽ tự động yêu cầu quyền Administrator khi cần
   - Hoặc click chuột phải và chọn "Run as Administrator"

### Lưu ý quan trọng
- Chương trình yêu cầu quyền Administrator để thực hiện các thao tác hệ thống
- Một số tính năng chỉ hoạt động trên Windows 10/11 Pro
- Cần kết nối internet để sử dụng tính năng cài đặt phần mềm
- Chương trình đã được đóng gói đầy đủ, không cần cài đặt Python hay các thư viện khác

### Yêu cầu hệ thống
- Windows 10/11
- Quyền Administrator
- Kết nối internet (cho tính năng cài đặt phần mềm)

## Hướng dẫn cho nhà phát triển

### Cài đặt môi trường phát triển
1. Cài đặt Python 3.8 trở lên
2. Cài đặt PyInstaller: `pip install pyinstaller`

### Tạo file thực thi
1. Mở Command Prompt với quyền Administrator
2. Di chuyển đến thư mục dự án
3. Chạy lệnh: `pyinstaller --onefile --windowed install.py`

### Cấu trúc dự án
- `install.py`: Mã nguồn chính
- `dist/`: Thư mục chứa file thực thi
- `build/`: Thư mục tạm thời trong quá trình build
