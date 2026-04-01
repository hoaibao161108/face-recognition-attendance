import tkinter as tk
from tkinter import messagebox
from datetime import datetime, time
import cv2
from PIL import Image, ImageTk
import os
import pandas as pd
import numpy as np
import unicodedata
import subprocess
from cuocthiAI import FaceAnalyzer
import shutil
import time

EXCEL_PATH = r"D:\Thi_AI\PYTHON\danh_sach_cham_cong.xlsx"
INFO_EXCEL_PATH = r"D:\Thi_AI\PYTHON\images\thong_tin_nguoi_dung.xlsx"
IMAGE_DIR = r"D:\Thi_AI\PYTHON\images"

cap = None
running = False
after_id = None
face_analyzer = FaceAnalyzer()
camera_target_frame = None
camera_label = None
camera_status_label = None
name_entry = None
custom_time = None

def normalize_filename(name):
    name = unicodedata.normalize('NFD', name)
    name = ''.join(c for c in name if unicodedata.category(c) != 'Mn')
    name = name.replace(' ', '_').replace('đ', 'd').replace('Đ', 'D')
    return name.strip()

def clean_directory(directory):
    if os.path.exists(directory):
        for item in os.listdir(directory):
            item_path = os.path.join(directory, item)
            try:
                if os.path.isdir(item_path):
                    shutil.rmtree(item_path)
                elif os.path.isfile(item_path) and item.lower().endswith(('.jpg', '.jpeg', '.png')):
                    os.remove(item_path)
            except Exception as e:
                print(f"Lỗi khi xóa {item_path}: {e}")

def show_frame():
    global cap, running, camera_label, after_id
    if not running or not cap or not cap.isOpened():
        if camera_status_label and camera_status_label.winfo_exists():
            camera_status_label.config(text="Camera: Tắt", fg="red")
        status_label.config(text="Camera: Tắt", fg="red")
        return
    
    try:
        ret, frame = cap.read()
        if not ret:
            status_label.config(text="Lỗi: Không đọc được frame", fg="red")
            if camera_status_label and camera_status_label.winfo_exists():
                camera_status_label.config(text="Lỗi: Không đọc được frame", fg="red")
            messagebox.showerror("Lỗi", "Không thể đọc frame từ camera.")
            tat_camera()
            return
        
        frame = face_analyzer.process_frame(frame, custom_time)
        img = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        img = cv2.resize(img, (640, 640))
        img = Image.fromarray(img)
        imgtk = ImageTk.PhotoImage(image=img)
        
        if camera_label and camera_label.winfo_exists():
            camera_label.imgtk = imgtk
            camera_label.configure(image=imgtk)
            
        status_label.config(text="Camera: Đang chạy", fg="green")
        if camera_status_label and camera_status_label.winfo_exists():
            camera_status_label.config(text="Camera: Đang chạy", fg="green")
            
    except Exception as e:
        print(f"Lỗi xử lý frame: {e}")
        status_label.config(text="Lỗi xử lý frame", fg="red")
        tat_camera()
        return
    
    after_id = root.after(10, show_frame)

def delayed_capture(delay_seconds, message):
    """Hàm trợ giúp để chụp ảnh với độ trễ"""
    for i in range(delay_seconds, 0, -1):
        status_label.config(text=f"{message}... {i} giây")
        root.update()
        time.sleep(1)
    return cap.read()

def bat_camera():
    global cap, running, camera_label, camera_target_frame
    if running:
        return

    hien_thi_camera()

    try:
        status_label.config(text="Đang mở camera...", fg="orange")
        clean_directory(IMAGE_DIR)
        face_analyzer.is_saving_user = False  # Đặt lại cờ khi bật camera

        for index in range(3):
            cap = cv2.VideoCapture(index)
            if cap.isOpened():
                running = True
                if camera_target_frame and not camera_label:
                    camera_label = tk.Label(camera_target_frame, bg="black")
                    camera_label.pack(expand=True, fill="both")
                show_frame()
                messagebox.showinfo("Thông báo", "🟢 Camera đã bật.")
                return
            time.sleep(1)

        status_label.config(text="Lỗi: Không mở được camera", fg="red")
        messagebox.showerror("Lỗi", "❌ Không tìm thấy camera.")
        
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể bật camera.\n{e}")
        status_label.config(text="Camera lỗi", fg="red")

def tat_camera():
    global cap, running, camera_label, after_id
    if not running:
        return
        
    running = False
    if cap:
        cap.release()
        cap = None
        
    if after_id:
        root.after_cancel(after_id)
        after_id = None
        
    try:
        face_analyzer.export_logs()
        if not os.path.exists(r"D:\Thi_AI\PYTHON\backup"):
            os.makedirs(r"D:\Thi_AI\PYTHON\backup")
        shutil.copy(INFO_EXCEL_PATH, r"D:\Thi_AI\PYTHON\backup\thong_tin_nguoi_dung_backup.xlsx")
    except Exception as e:
        print(f"Lỗi khi sao lưu: {e}")
        
    if camera_status_label and camera_status_label.winfo_exists():
        camera_status_label.config(text="Camera: Tắt", fg="red")
    status_label.config(text="Camera: Tắt", fg="red")
    messagebox.showinfo("Thông báo", "🔴 Camera đã tắt & lưu log.")
    
def show_set_time_dialog():
    global custom_time
    popup = tk.Toplevel(root)
    popup.title("Chỉnh thời gian")
    popup.geometry("300x200")
    popup.configure(bg="white")
    popup.resizable(False, False)

    tk.Label(popup, text="Nhập thời gian (HH:MM:SS):", bg="white", font=("Arial", 12)).pack(pady=10)
    time_entry = tk.Entry(popup, font=("Arial", 12), width=20)
    time_entry.pack(pady=5)

    def apply_time():
        global custom_time
        try:
            time_str = time_entry.get().strip()
            datetime.strptime(time_str, "%H:%M:%S")
            custom_time = time_str
            face_analyzer.set_deadline(custom_time)
            messagebox.showinfo("Thành công", f"Thời gian đã được đặt: {custom_time}")
            popup.destroy()
            
            if running:
                root.after_cancel(after_id)
                show_frame()
                
            for widget in content_frame.winfo_children():
                if isinstance(widget, tk.Label) and "Thời gian hiện tại" in widget.cget("text"):
                    widget.config(text=f"🕒 Thời gian hiện tại: {custom_time}")
        except ValueError:
            messagebox.showerror("Lỗi", "Định dạng thời gian không hợp lệ!")

    tk.Button(popup, text="Áp dụng", font=("Arial", 12), bg="#2ecc71", fg="white", command=apply_time).pack(pady=20)

def hien_thi_form_nhap_thong_tin():
    global entries, popup
    if not running:
        bat_camera()
        
    popup = tk.Toplevel(root)
    popup.title("Nhập thông tin")
    popup.geometry("420x480")
    popup.configure(bg="white")
    popup.resizable(False, False)

    fields = ["Họ tên", "Mã số", "Lớp/Phòng ban", "Số điện thoại", "Email"]
    entries = []

    for field in fields:
        tk.Label(popup, text=f"{field}:", bg="white", font=("Arial", 12)).pack(pady=(10, 2))
        entry = tk.Entry(popup, font=("Arial", 12), width=30)
        entry.pack(pady=2)
        entries.append(entry)

    tk.Button(popup, text="💾 Lưu", font=("Arial", 12), bg="#2ecc71", fg="white", command=luu_nguoi_dung).pack(pady=20)

def luu_nguoi_dung():
    global entries
    values = [e.get().strip() for e in entries]
    if any(not v for v in values):
        messagebox.showerror("Lỗi", "Vui lòng nhập đầy đủ thông tin!")
        return

    name, mssv, lop, sdt, email = values
    normalized_name = normalize_filename(name)

    if not cap or not cap.isOpened():
        messagebox.showerror("Lỗi", "Camera chưa bật!")
        return

    # Đặt cờ đang lưu thông tin
    face_analyzer.is_saving_user = True
    person_dir = os.path.join(IMAGE_DIR, normalized_name)
    
    try:
        os.makedirs(person_dir, exist_ok=True)
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể tạo thư mục: {e}")
        face_analyzer.is_saving_user = False
        return

    # Chụp 3 góc ảnh với thời gian chờ mới (4 giây mỗi giai đoạn)
    angles = ["front", "left", "right"]
    image_paths = []
    
    for angle in angles:
        # Hiển thị hướng dẫn
        if angle == "left":
            messagebox.showinfo("Hướng dẫn", "Vui lòng nghiêng đầu sang TRÁI 30 độ")
        elif angle == "right":
            messagebox.showinfo("Hướng dẫn", "Vui lòng nghiêng đầu sang PHẢI 30 độ")
        
        # Chờ 4 giây sau thông báo
        for i in range(4, 0, -1):
            status_label.config(text=f"Chuẩn bị chụp ảnh {angle}... {i} giây")
            root.update()
            time.sleep(1)
        
        # Chụp ảnh sau 4 giây chờ
        ret, frame = cap.read()
        if not ret:
            messagebox.showerror("Lỗi", "Không thể chụp ảnh từ camera!")
            face_analyzer.is_saving_user = False
            return
            
        # Lưu ảnh
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        img_path = os.path.join(person_dir, f"{angle}_{timestamp}.jpg")
        
        try:
            if cv2.imwrite(img_path, frame):
                image_paths.append(img_path)
                face_analyzer.update_face_database(name, img_path)
        except Exception as e:
            print(f"Lỗi lưu ảnh {angle}: {e}")

    # Kết thúc quá trình lưu
    face_analyzer.is_saving_user = False
    status_label.config(text="Camera: Đang chạy", fg="green")
    
    if not image_paths:
        messagebox.showerror("Lỗi", "Không lưu được ảnh nào!")
        return

    # Lưu thông tin vào Excel
    now = datetime.now()
    time_str = custom_time if custom_time else now.strftime("%H:%M:%S")
    deadline = face_analyzer.get_deadline()
    status = "Đúng giờ" if datetime.strptime(time_str, "%H:%M:%S").time() <= deadline else "Trễ giờ"
    
    try:
        late_minutes = max(0, (datetime.strptime(time_str, "%H:%M:%S") - 
                          datetime.strptime(face_analyzer.custom_deadline, "%H:%M:%S")).seconds // 60)
    except:
        late_minutes = 0

    new_data = pd.DataFrame([{
        "Họ tên": name,
        "Mã số": mssv,
        "Lớp/Phòng ban": lop,
        "Số điện thoại": sdt,
        "Email": email,
        "Giờ cần có mặt": face_analyzer.custom_deadline,
        "Thời gian điểm danh": time_str,
        "Trạng thái": status,
        "Đi trễ (phút)": late_minutes,
        "Ngày thêm": now.strftime("%d/%m/%Y")
    }])

    try:
        if os.path.exists(INFO_EXCEL_PATH):
            old_data = pd.read_excel(INFO_EXCEL_PATH, engine='openpyxl')
            df_all = pd.concat([old_data, new_data], ignore_index=True)
            df_all = df_all.drop_duplicates(subset=["Mã số"], keep="last")
        else:
            df_all = new_data
            
        df_all.to_excel(INFO_EXCEL_PATH, index=False, engine='openpyxl')
        messagebox.showinfo("Thành công", "Đã lưu thông tin thành công!")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể lưu Excel: {e}")
    finally:
        popup.destroy()

def show_delete_user_dialog():
    popup = tk.Toplevel(root)
    popup.title("Xóa thông tin")
    popup.geometry("300x200")
    popup.configure(bg="white")
    popup.resizable(False, False)

    tk.Label(popup, text="Nhập mã số:", bg="white", font=("Arial", 12)).pack(pady=10)
    mssv_entry = tk.Entry(popup, font=("Arial", 12), width=20)
    mssv_entry.pack(pady=5)

    def delete_user():
        mssv = mssv_entry.get().strip()
        if not mssv:
            messagebox.showerror("Lỗi", "Vui lòng nhập mã số!")
            return
            
        try:
            if not os.path.exists(INFO_EXCEL_PATH):
                messagebox.showerror("Lỗi", "File thông tin không tồn tại!")
                popup.destroy()
                return
                
            df = pd.read_excel(INFO_EXCEL_PATH, engine='openpyxl')
            df["Mã số"] = df["Mã số"].astype(str).str.strip()
            
            if mssv not in df["Mã số"].values:
                messagebox.showerror("Lỗi", f"Mã số {mssv} không tồn tại!")
                popup.destroy()
                return
                
            # Xóa thông tin
            user_row = df[df["Mã số"] == mssv]
            name = user_row["Họ tên"].iloc[0]
            normalized_name = normalize_filename(name)
            person_dir = os.path.join(IMAGE_DIR, normalized_name)
            
            if os.path.exists(person_dir):
                shutil.rmtree(person_dir)
                
            df = df[df["Mã số"] != mssv]
            df.to_excel(INFO_EXCEL_PATH, index=False, engine='openpyxl')
            
            # Cập nhật database nhận diện
            face_analyzer.dataframe = face_analyzer.dataframe[face_analyzer.dataframe["Name"] != name]
            if len(face_analyzer.dataframe) > 0:
                face_analyzer.embeddings = np.stack(face_analyzer.dataframe["embedding"].values)
            else:
                face_analyzer.embeddings = np.array([])
                
            messagebox.showinfo("Thành công", f"Đã xóa thông tin {name}!")
            popup.destroy()
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xóa: {e}")
            popup.destroy()

    tk.Button(popup, text="Xóa", font=("Arial", 12), bg="#c0392b", fg="white", command=delete_user).pack(pady=20)

def hien_thi_camera():
    global camera_target_frame, camera_label, camera_status_label
    for widget in content_frame.winfo_children():
        widget.destroy()
        
    camera_target_frame = tk.Frame(content_frame, bg="#ecf0f1", bd=2, relief="groove")
    camera_target_frame.pack(pady=10, expand=True, fill="both")
    
    camera_label = tk.Label(camera_target_frame, text="📷 Camera sẽ hiển thị ở đây",
                          font=("Arial", 14), width=60, height=15, bg="black", fg="white")
    camera_label.pack(expand=True, fill="both")
    
    input_frame = tk.Frame(content_frame, bg="#ecf0f1")
    input_frame.pack(pady=5)
    
    tk.Button(input_frame, text="💾 Lưu thông tin", font=("Arial", 12),
             bg="#2ecc71", fg="white", command=hien_thi_form_nhap_thong_tin).pack(side="left", padx=5)
             
    tk.Button(input_frame, text="🕒 Chỉnh giờ", font=("Arial", 12),
             bg="#3498db", fg="white", command=show_set_time_dialog).pack(side="left", padx=5)
             
    camera_status_label = tk.Label(content_frame, text="Camera: Tắt", font=("Arial", 14, "bold"),
                                 bg="#ecf0f1", fg="red")
    camera_status_label.pack(pady=5)
    
    clock_label = tk.Label(content_frame, font=("Arial", 16), bg="#ecf0f1", fg="#2c3e50")
    clock_label.pack()
    
    def update_clock():
        now = custom_time if custom_time else datetime.now().strftime("%H:%M:%S")
        clock_label.config(text=f"🕒 Thời gian hiện tại: {now}")
        clock_label.after(1000, update_clock)
        
    update_clock()
    
    if running:
        camera_label.destroy()
        camera_label = tk.Label(camera_target_frame, bg="black")
        camera_label.pack(expand=True, fill="both")
        show_frame()

def open_excel_file():
    try:
        if os.path.exists(INFO_EXCEL_PATH):
            os.startfile(INFO_EXCEL_PATH)
        else:
            messagebox.showerror("Lỗi", "File Excel không tồn tại!")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể mở file: {e}")

def hien_thi_trang_chu():
    global camera_target_frame
    for widget in content_frame.winfo_children():
        widget.destroy()
        
    button_frame = tk.Frame(content_frame, bg="#ecf0f1")
    button_frame.pack(expand=True)
    
    buttons = [
        ("▶ Bật Camera", "#27ae60", bat_camera),
        ("⏹ Tắt Camera", "#c0392b", tat_camera),
        ("🗑 Xóa thông tin", "#8e44ad", show_delete_user_dialog),
        ("📂 Mở Excel", "#3498db", open_excel_file)
    ]
    
    for i, (text, color, cmd) in enumerate(buttons):
        row = i // 2
        col = i % 2
        tk.Button(button_frame, text=text, font=("Arial", 22, "bold"),
                bg=color, fg="white", command=cmd, width=22, height=4).grid(row=row, column=col, padx=30, pady=30)

def dang_xuat():
    if messagebox.askokcancel("Thoát", "Bạn có chắc muốn thoát?"):
        tat_camera()
        root.destroy()

# Tạo giao diện chính
root = tk.Tk()
root.title("Hệ thống điểm danh bằng khuôn mặt")
root.geometry("1200x700")
root.configure(bg="#f0f0f0")
root.resizable(True, True)

# Header
header_frame = tk.Frame(root, bg="#2c3e50", height=60)
header_frame.pack(side="top", fill="x")

tk.Label(header_frame, text="🎓 HỆ THỐNG ĐIỂM DANH BẰNG KHUÔN MẶT",
        font=("Arial", 20, "bold"), fg="white", bg="#2c3e50").pack(pady=10)

status_label = tk.Label(root, text="", font=("Arial", 12), bg="#f0f0f0", fg="black")
status_label.pack()

# Main container
main_container = tk.Frame(root, bg="#f0f0f0")
main_container.pack(fill="both", expand=True)

# Menu
menu_frame = tk.Frame(main_container, bg="#34495e", width=200)
menu_frame.pack(side="left", fill="y")

tk.Label(menu_frame, text="📁 MENU", font=("Arial", 14, "bold"),
        fg="white", bg="#34495e", pady=10).pack()

menu_items = [
    ("Trang chủ", hien_thi_trang_chu),
    ("Camera", hien_thi_camera),
    ("Đăng xuất", dang_xuat)
]

for name, cmd in menu_items:
    tk.Button(menu_frame, text=name, font=("Arial", 12), fg="white",
            bg="#34495e", relief="flat", activebackground="#2c3e50", 
            anchor="w", command=cmd).pack(fill="x", padx=10, pady=5)

# Content
content_frame = tk.Frame(main_container, bg="#ecf0f1")
content_frame.pack(side="right", fill="both", expand=True)

hien_thi_trang_chu()
root.mainloop()