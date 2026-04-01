# 🎓 Hệ Thống Điểm Danh Bằng Nhận Diện Khuôn Mặt

Một ứng dụng Desktop hỗ trợ điểm danh thông minh, tự động nhận diện khuôn mặt người dùng qua camera, đối chiếu dữ liệu, ghi nhận thời gian check-in và tính toán thời gian đi trễ.

## 🚀 Tính năng nổi bật
* **Nhận diện thời gian thực:** Phát hiện và trích xuất vector đặc trưng khuôn mặt với độ chính xác cao bằng mô hình Deep Learning (InsightFace).
* **Giao diện trực quan:** Giao diện người dùng (GUI) thân thiện được xây dựng bằng Tkinter, hiển thị luồng video và luồng sự kiện theo thời gian thực.
* **Quản lý dữ liệu:** Lưu vết lịch sử điểm danh và quản lý thông tin nhân sự.
* **Tự động hóa:** Tích hợp tính năng gửi email báo cáo trạng thái check-in (SMTP).

## 🛠️ Công nghệ sử dụng
* **Ngôn ngữ cốt lõi:** Python 3
* **AI & Computer Vision:** OpenCV, InsightFace, ONNX Runtime
* **Xây dựng Giao diện:** Tkinter
* **Xử lý Dữ liệu:** Pandas, NumPy, Scikit-learn

## ⚙️ Hướng dẫn cài đặt
**Bước 1:** Clone repository này về máy cục bộ.
**Bước 2:** Cài đặt các thư viện môi trường cần thiết:
   `pip install -r requirements.txt`
**Bước 3:** Bổ sung Model AI.
   *Do dung lượng giới hạn của GitHub, các mô hình học sâu đã được loại bỏ khỏi mã nguồn.*
   Tải model `buffalo_sc` (định dạng .onnx) và giải nén vào đường dẫn: `insightface_model/models/buffalo_sc/`.
**Bước 4:** Khởi chạy hệ thống bằng lệnh:
   `python giao_dien.py`
