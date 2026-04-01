import os
import cv2
import insightface
import numpy as np
import pandas as pd
from sklearn.metrics import pairwise
from datetime import datetime
from PIL import ImageFont, ImageDraw, Image
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
import threading

IMAGE_DIR = r'D:\Thi_AI\PYTHON\images'
EXCEL_PATH = r"D:\Thi_AI\PYTHON\images\thong_tin_nguoi_dung.xlsx"
CHEAT_DIR = r"D:\Thi_AI\PYTHON\gianlan"

class FaceAnalyzer:
    def __init__(self, thresh=0.15):
        self.faceapp = insightface.app.FaceAnalysis(
            name='buffalo_sc',
            root='insightface_model',
            providers=['CPUExecutionProvider']
        )
        self.faceapp.prepare(ctx_id=0, det_size=(320, 320))
        self.dataframe, self.embeddings = self.load_face_database()
        self.thresh = thresh
        self.recognized_log = []
        self.custom_deadline = "06:15:00"
        self.prev_log_hash = None
        self.sender_email = "datdcts02142@gmail.com"
        self.sender_password = "eyxs yyqa qwwq sjab"
        self.receiver_email = "datdcts02142@gmail.com"
        self.prev_bboxes = {}
        self.is_saving_user = False  # Thêm cờ để kiểm tra trạng thái lưu người dùng

    def set_deadline(self, deadline):
        try:
            datetime.strptime(deadline, "%H:%M:%S")
            self.custom_deadline = deadline
            self.update_excel_deadline()
            self.update_log_status(deadline)
        except ValueError:
            self.custom_deadline = "06:15:00"

    def get_deadline(self):
        try:
            return datetime.strptime(self.custom_deadline, "%H:%M:%S").time()
        except ValueError:
            return datetime.strptime("06:15:00", "%H:%M:%S").time()

    def draw_text_unicode(self, img_cv2, text, position, font_path="C:/Windows/Fonts/arial.ttf", font_size=24, color=(0, 255, 0)):
        img_pil = Image.fromarray(cv2.cvtColor(img_cv2, cv2.COLOR_BGR2RGB))
        draw = ImageDraw.Draw(img_pil)
        font = ImageFont.truetype(font_path, font_size)
        draw.text(position, text, font=font, fill=color[::-1])
        return cv2.cvtColor(np.array(img_pil), cv2.COLOR_RGB2BGR)

    def list_visible_dirs(self, directory):
        return [f for f in os.listdir(directory) if not f.startswith('.') and os.path.isdir(os.path.join(directory, f))]

    def list_visible_images_inpath(self, directory):
        return [f for f in os.listdir(directory) if not f.startswith('.') and f.lower().endswith(('.jpg', '.jpeg', '.png'))]

    def load_face_database(self):
        all_info = []
        embeddings_list = []
        names = []
        folders = self.list_visible_dirs(IMAGE_DIR)
        for folder in folders:
            name = folder
            folder_path = os.path.join(IMAGE_DIR, folder)
            img_files = self.list_visible_images_inpath(folder_path)
            for file in img_files:
                img_path = os.path.join(folder_path, file)
                img_arr = cv2.imread(img_path)
                if img_arr is None:
                    continue
                result = self.faceapp.get(img_arr, max_num=1)
                if len(result) > 0:
                    embedding = result[0]['embedding']
                    all_info.append([name, embedding])
                    embeddings_list.append(embedding)
                    names.append(name)
        df = pd.DataFrame({"Name": names, "embedding": embeddings_list}) if all_info else pd.DataFrame(columns=["Name", "embedding"])
        return df, np.asarray(embeddings_list)

    def update_face_database(self, name, image_path):
        img_arr = cv2.imread(image_path)
        if img_arr is None:
            return
        result = self.faceapp.get(img_arr, max_num=1)
        if len(result) > 0:
            embedding = result[0]['embedding']
            new_row = pd.DataFrame({"Name": [name], "embedding": [embedding]})
            self.dataframe = pd.concat([self.dataframe, new_row], ignore_index=True)
            if len(self.embeddings) > 0:
                self.embeddings = np.vstack([self.embeddings, embedding])
            else:
                self.embeddings = np.array([embedding])

    def send_email(self, subject, body, image_path=None):
        """Gửi email với hoặc không có ảnh đính kèm"""
        msg = MIMEMultipart()
        msg['From'] = self.sender_email
        msg['To'] = self.receiver_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))
        
        if image_path and os.path.exists(image_path):
            with open(image_path, 'rb') as f:
                img = MIMEImage(f.read())
                img.add_header('Content-Disposition', 'attachment', filename=os.path.basename(image_path))
                msg.attach(img)
        
        try:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(self.sender_email, self.sender_password)
            server.send_message(msg)
            server.quit()
        except Exception as e:
            print(f"Lỗi gửi email: {e}")

    def send_email_in_thread(self, subject, body, image_path=None):
        """Gửi email trong một luồng riêng"""
        email_thread = threading.Thread(target=self.send_email, args=(subject, body, image_path))
        email_thread.daemon = True
        email_thread.start()

    def send_late_email(self, name, time, date):
        msg = MIMEMultipart()
        msg['From'] = self.sender_email
        msg['To'] = self.receiver_email
        msg['Subject'] = "Thông Báo Đi Trễ"
        body = f"Nhân viên {name} đã đi trễ vào lúc {time} ngày {date}."
        msg.attach(MIMEText(body, 'plain'))
        try:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(self.sender_email, self.sender_password)
            server.send_message(msg)
            server.quit()
        except:
            pass

    def update_log_status(self, custom_time):
        deadline = self.get_deadline()
        current_log_hash = hash(str(self.recognized_log))
        if self.prev_log_hash != current_log_hash:
            for log in self.recognized_log:
                cur_time = datetime.strptime(log['time'], "%H:%M:%S").time()
                log['status'] = "Đúng giờ" if cur_time <= deadline else "Trễ giờ"
                if log['status'] == "Trễ giờ" and not log.get('email_sent', False):
                    self.send_late_email(log['name'], log['time'], log['date'])
                    log['email_sent'] = True
            self.prev_log_hash = current_log_hash

    def update_excel_deadline(self):
        try:
            if os.path.exists(EXCEL_PATH):
                df = pd.read_excel(EXCEL_PATH, engine='openpyxl')
                
                if "Giờ cần có mặt" in df.columns:
                    df["Giờ cần có mặt"] = self.custom_deadline
                elif "Giờ cán cô mặt" in df.columns:
                    df = df.rename(columns={"Giờ cán cô mặt": "Giờ cần có mặt"})
                    df["Giờ cần có mặt"] = self.custom_deadline
                else:
                    df["Giờ cần có mặt"] = self.custom_deadline

                deadline_time = datetime.strptime(self.custom_deadline, "%H:%M:%S").time()
                fmt = "%H:%M:%S"
                for index, row in df.iterrows():
                    try:
                        checkin_time = datetime.strptime(row["Thời gian điểm danh"], fmt).time()
                        status = "Đúng giờ" if checkin_time <= deadline_time else "Trễ giờ"
                        df.at[index, "Trạng thái"] = status
                        
                        checkin_dt = datetime.strptime(row["Thời gian điểm danh"], fmt)
                        deadline_dt = datetime.strptime(self.custom_deadline, fmt)
                        late_minutes = max(0, int((checkin_dt - deadline_dt).total_seconds() / 60))
                        df.at[index, "Đi trễ (phút)"] = late_minutes
                    except Exception as e:
                        print(f"Lỗi khi cập nhật trạng thái cho bản ghi {index}: {e}")
                        df.at[index, "Trạng thái"] = "Không xác định"
                        df.at[index, "Đi trễ (phút)"] = 0

                df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            else:
                print(f"File Excel {EXCEL_PATH} không tồn tại.")
        except Exception as e:
            print(f"Lỗi update_excel_deadline: {e}")

    def check_detailed_cheating(self, frame, face):
        """Kiểm tra gian lận chi tiết và gửi email với ảnh nếu phát hiện"""
        try:
            emb = face['embedding'].reshape(1, -1)
            name_key = hash(str(emb.tobytes()))
            
            if len(self.embeddings) > 0:
                cosine_similar = pairwise.cosine_similarity(self.embeddings, emb)
                df_cosine = self.dataframe.copy()
                df_cosine['cosine'] = cosine_similar.flatten()
                df_filtered = df_cosine.query(f'cosine > {self.thresh}').reset_index(drop=True)
                
                if len(df_filtered) > 0:
                    best_idx = df_filtered['cosine'].idxmax()
                    name = df_filtered.loc[best_idx]['Name']
                    
                    landmarks = face['kps']
                    left_eye = landmarks[0]
                    right_eye = landmarks[1]
                    nose = landmarks[2]
                    
                    eye_vector = right_eye - left_eye
                    nose_vector = nose - left_eye
                    angle = np.degrees(np.arctan2(nose_vector[1], nose_vector[0]) - np.arctan2(eye_vector[1], eye_vector[0]))
                    angle = abs(angle)
                    
                    if angle > 45:
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        cheat_path = os.path.join(CHEAT_DIR, f"{name}_cheat_{timestamp}.jpg")
                        if not os.path.exists(CHEAT_DIR):
                            os.makedirs(CHEAT_DIR)
                        cv2.imwrite(cheat_path, frame)
                        self.send_email_in_thread(
                            subject="Cảnh Báo Gian Lận",
                            body=f"Phát hiện gian lận từ nhân viên {name} vào lúc {timestamp}.",
                            image_path=cheat_path
                        )
                        return True, name
            return False, None
        except Exception as e:
            print(f"Lỗi khi kiểm tra gian lận: {e}")
            return False, None

    def process_frame(self, frame, custom_time=None):
        try:
            faces = self.faceapp.get(frame, max_num=0)
            num_known = 0
            num_unknown = 0
            recognized_names = set()

            for face in faces:
                bbox = face['bbox'].astype(int)
                emb = face['embedding'].reshape(1, -1)
                name_key = hash(str(emb.tobytes()))

                if name_key in self.prev_bboxes:
                    prev_bbox = self.prev_bboxes[name_key]
                    bbox = (prev_bbox * 0.6 + bbox * 0.4).astype(int)
                self.prev_bboxes[name_key] = bbox

                # Bỏ qua kiểm tra gian lận nếu đang lưu thông tin người dùng
                if not self.is_saving_user:
                    is_cheating, cheat_name = self.check_detailed_cheating(frame, face)
                    if is_cheating:
                        label = f"{cheat_name} (CẢNH BÁO GIAN LẬN)"
                        color = (0, 0, 255)
                        x1, y1, x2, y2 = bbox
                        frame = cv2.rectangle(frame, (x1, y1), (x2, y2), color, 2)
                        frame = self.draw_text_unicode(frame, label, (x1, y1 - 25), font_size=20, color=color)
                        continue

                if len(self.embeddings) > 0:
                    cosine_similar = pairwise.cosine_similarity(self.embeddings, emb)
                    df_cosine = self.dataframe.copy()
                    df_cosine['cosine'] = cosine_similar.flatten()
                    df_filtered = df_cosine.query(f'cosine > {self.thresh}').reset_index(drop=True)
                    if len(df_filtered) > 0:
                        best_idx = df_filtered['cosine'].idxmax()
                        name = df_filtered.loc[best_idx]['Name']
                        cosine_score = f"{df_filtered.loc[best_idx]['cosine']:.2f} ({len(df_filtered)})"
                        if name not in recognized_names:
                            num_known += 1
                            recognized_names.add(name)
                        now = datetime.now()
                        recognition_time = now.strftime("%H:%M:%S")
                        cur_date = now.strftime("%d/%m/%Y")
                        deadline = self.get_deadline()
                        status = "Đúng giờ" if datetime.strptime(recognition_time, "%H:%M:%S").time() <= deadline else "Trễ giờ"
                        found = False
                        for log in self.recognized_log:
                            if log['name'] == name and log['date'] == cur_date:
                                log['time'] = recognition_time
                                log['status'] = status
                                found = True
                                break
                        if not found:
                            self.recognized_log.append({
                                'name': name,
                                'time': recognition_time,
                                'date': cur_date,
                                'status': status,
                                'email_sent': False
                            })
                        self.update_log_status(recognition_time)
                        label = f"{name} ({status}) {cosine_score}"
                        color = (0, 255, 0)
                    else:
                        num_unknown += 1
                        label = "Khuôn mặt không xác định"
                        color = (255, 255, 255)
                else:
                    num_unknown += 1
                    label = "Khuôn mặt không xác định"
                    color = (255, 255, 255)

                x1, y1, x2, y2 = bbox
                frame = cv2.rectangle(frame, (x1, y1), (x2, y2), color, 2)
                frame = self.draw_text_unicode(frame, label, (x1, y1 - 25), font_size=20, color=color)

            frame = self.draw_text_unicode(frame, f'Đã nhận: {num_known}', (10, 30), font_size=26, color=(255, 255, 0))
            frame = self.draw_text_unicode(frame, f'Không rõ: {num_unknown}', (10, 65), font_size=26, color=(255, 255, 255))
            return frame
        except Exception as e:
            print(f"Lỗi trong process_frame: {e}")
            return frame