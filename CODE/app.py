import sys
import cv2
import face_recognition
import os
import numpy as np
from datetime import datetime
import openpyxl 
#GUI
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import DateEntry
from PIL import Image, ImageTk
#Sms
from twilio.rest import Client

import pandas as pd
#Gmail
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

if getattr(sys, 'frozen', False):
    # Đường dẫn tới thư mục chứa executable
    application_path = sys._MEIPASS
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

class CameraApp:
    def __init__(self):
        self.window = ctk.CTk()
        self.window.title("HỆ THỐNG ĐIỂM DANH THÔNG QUA CAMERA")
        self.window.geometry("1500x850")
        self.window.resizable(False, False)

        self.video_capture = None
        self.encoded_known_faces = []
        self.known_names = []
        self.dataframe = None
        self.last_attendance_time = {}
        self.attendance_cooldown = 3
        self.excel_data = None
        
        #Đường dẫn thư mục ảnh mặc định
        self.default_faces_directory = r"D:\python\DATA"
        
        #Tự động tải ảnh khuôn mặt khi khởi động
        self.load_known_faces(self.default_faces_directory)
        
        # Set custom appearance
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        
        # Create GUI
        self.create_gui()
        
        # Start update loop
        self.update_loop()

            
    def create_gui(self):
        # Title Label with large font
        title_label = ctk.CTkLabel(
            self.window,
            text="HỆ THỐNG ĐIỂM DANH THÔNG QUA CAMERA",
            font=("Helvetica", 24, "bold")
        )
        title_label.pack(pady=20)

        # Main content frame
        content_frame = ctk.CTkFrame(self.window)
        content_frame.pack(expand=True, fill="both", padx=20, pady=20)

        # Left section - Camera
        camera_section = ctk.CTkFrame(content_frame)
        camera_section.pack(side="left", padx=10, pady=10, fill="both")

        camera_label = ctk.CTkLabel(
            camera_section,
            text="Hệ thống CAMERA",
            font=("Helvetica", 16, "bold")
        )
        camera_label.pack(pady=10)

        # Camera Canvas
        self.camera_canvas = ctk.CTkCanvas(
            camera_section,
            width=640,
            height=480,
            bg="lightgray"
        )
        self.camera_canvas.pack(padx=10, pady=10)

        # Camera controls
        camera_controls = ctk.CTkFrame(camera_section)
        camera_controls.pack(fill="x", padx=10, pady=10)

        id_label = ctk.CTkLabel(camera_controls, text="ID")
        id_label.pack(side="left", padx=5)

        self.id_entry = ctk.CTkEntry(camera_controls, width=100)
        self.id_entry.pack(side="left", padx=5)

        self.camera_button = ctk.CTkButton(
            camera_controls,
            text="Bật Camera",
            command=self.toggle_camera,
            width=120
        )
        self.camera_button.pack(side="left", padx=5)

        # Right section
        right_section = ctk.CTkFrame(content_frame)
        right_section.pack(side="left", padx=10, pady=10, fill="both", expand=True)

        # Excel file selection
        excel_frame = ctk.CTkFrame(right_section)
        excel_frame.pack(fill="x", pady=10)

        excel_label = ctk.CTkLabel(
            excel_frame,
            text="File Excel đã chọn",
            font=("Helvetica", 14)
        )
        excel_label.pack(pady=5)

        excel_controls = ctk.CTkFrame(excel_frame)
        excel_controls.pack(fill="x")

        self.excel_entry = ctk.CTkEntry(
            excel_controls,
            placeholder_text="Chưa chọn file Excel",
            width=300
        )
        self.excel_entry.pack(side="left", padx=5)

        excel_button = ctk.CTkButton(
            excel_controls,
            text="Chọn file Excel",
            command=self.select_excel_file,
            width=120
        )
        excel_button.pack(side="left", padx=5)

        display_excel_button = ctk.CTkButton(
            excel_controls,
            text="Hiển thị file Excel",
            command=self.display_excel_data,
            width=120
        )
        display_excel_button.pack(side="left", padx=5)
        
        # button send email to parents
        send_email_button = ctk.CTkButton(
            excel_controls,
            text="Gửi email cho phụ huynh",
            command=self.send_email_to_parents,
            width=180
        )
        send_email_button.pack(side="left", padx=5)

        # Date selection
        date_frame = ctk.CTkFrame(right_section)
        date_frame.pack(fill="x", pady=10)

        date_label = ctk.CTkLabel(
            date_frame,
            text="Ngày điểm danh",
            font=("Helvetica", 14)
        )
        date_label.pack(pady=5)

        date_controls = ctk.CTkFrame(date_frame)
        date_controls.pack(fill="x")

        self.date_entry = DateEntry(
            date_controls,
            width=30,
            background='darkblue',
            foreground='white',
            borderwidth=2,
            date_pattern='dd/mm/yyyy'
        )
        self.date_entry.pack(side="left", padx=5)

        date_button = ctk.CTkButton(
            date_controls,
            text="Chọn ngày",
            width=120
        )
        date_button.pack(side="left", padx=5)

        # Create Treeview frame
        treeview_frame = ctk.CTkFrame(right_section)
        treeview_frame.pack(fill="both", expand=True, pady=10)

        # Create Treeview
        self.tree = ttk.Treeview(treeview_frame, selectmode='browse')
        self.tree.pack(fill="both", expand=True, padx=5, pady=5)

        # Add scrollbar to treeview
        scrollbar = ttk.Scrollbar(treeview_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Horizontal scrollbar
        h_scrollbar = ttk.Scrollbar(treeview_frame, orient="horizontal", command=self.tree.xview)
        h_scrollbar.pack(side="bottom", fill="x")
        self.tree.configure(xscrollcommand=h_scrollbar.set)

        # Style configuration for Treeview
        style = ttk.Style()
        style.configure("Treeview",
                       background="white",
                       foreground="black",
                       rowheight=25,
                       fieldbackground="white")
        style.configure("Treeview.Heading",
                       background="lightgray",
                       foreground="black",
                       relief="raised")

        # Buttons frame
        buttons_frame = ctk.CTkFrame(right_section)
        buttons_frame.pack(fill="x", pady=10, padx=10)

        unmarked_button = ctk.CTkButton(
            buttons_frame,
            text="Danh sách học sinh\nchưa điểm danh",
            command=self.show_unmarked,
            height=60
        )
        unmarked_button.pack(side="left", padx=5, expand=True)

        marked_button = ctk.CTkButton(
            buttons_frame,
            text="Danh sách học sinh\nđã điểm danh",
            command=self.show_marked,
            height=60
        )
        marked_button.pack(side="left", padx=5, expand=True)

        clear_button = ctk.CTkButton(
            buttons_frame,
            text="Xóa nội dung",
            command=self.clear_log,
            height=60
        )
        clear_button.pack(side="left", padx=5, expand=True)

    # def load_known_faces(self, directory_path):
    #     """
    #     Load known faces from a directory
    #     Filename is used as the student ID/name
    #     """
    #     self.encoded_known_faces = []
    #     self.known_names = []
        
    #     # Kiểm tra và thông báo nếu thư mục không tồn tại
    #     if not os.path.exists(directory_path):
    #         messagebox.showwarning("Cảnh báo", f"Thư mục {directory_path} không tồn tại.")
    #         return
        
    #     # Đếm số lượng ảnh được tải thành công
    #     successful_loads = 0
    #     failed_loads = 0
        
    #     # Iterate through image files in the directory
    #     for filename in os.listdir(directory_path):
    #         # Support common image formats
    #         if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif')):
    #             try:
    #                 # Remove file extension to get student ID/name
    #                 student_name = os.path.splitext(filename)[0]
                    
    #                 # Full path to the image
    #                 image_path = os.path.join(directory_path, filename)
                    
    #                 # Load image
    #                 image = face_recognition.load_image_file(image_path)
                    
    #                 # Encode face
    #                 face_encodings = face_recognition.face_encodings(image)
                    
    #                 # If a face is detected in the image
    #                 if face_encodings:
    #                     self.encoded_known_faces.append(face_encodings[0])
    #                     self.known_names.append(student_name)
    #                     successful_loads += 1
    #                 else:
    #                     failed_loads += 1
    #                     print(f"Không phát hiện khuôn mặt trong {filename}")
                
    #             except Exception as e:
    #                 failed_loads += 1
    #                 print(f"Lỗi xử lý {filename}: {e}")
        
    #     # Thông báo kết quả tải
    #     if successful_loads > 0:
    #         messagebox.showinfo("Thành công", 
    #                             f"Đã tải {successful_loads} khuôn mặt.\n"
    #                             f"Số ảnh không thành công: {failed_loads}")
    #     else:
    #         messagebox.showwarning("Cảnh báo", 
    #                                 f"Không tải được ảnh nào từ thư mục {directory_path}.\n"
    #                                 "Vui lòng kiểm tra lại thư mục.")

    # def mark_attendance(self, name):
    #     # Kiểm tra điều kiện ban đầu
    #     if not self.excel_file:
    #         return
        
    #     # Nếu chưa chọn ngày, sử dụng ngày hiện tại
    #     if not self.selected_date:
    #         self.selected_date = datetime.now().strftime("%Y-%m-%d")

    #     try:
    #         # Mở workbook Excel
    #         workbook = openpyxl.load_workbook(self.excel_file)
    #         sheet = workbook.active

    #         # Tìm cột cho ngày điểm danh
    #         column_letter = None
    #         for col in range(1, sheet.max_column + 1):
    #             cell_value = sheet.cell(row=1, column=col).value
    #             if cell_value == self.selected_date:
    #                 column_letter = col
    #                 break

    #         # Nếu chưa có cột cho ngày, tạo cột mới
    #         if column_letter is None:
    #             column_letter = sheet.max_column + 1
    #             sheet.cell(row=1, column=column_letter).value = self.selected_date
    #             print(f"Thêm cột mới cho ngày: {self.selected_date}")

    #         # Lấy thời gian điểm danh
    #         now = datetime.now()
    #         datetime_string = now.strftime("%H:%M:%S")

    #         # Biến để theo dõi trạng thái điểm danh
    #         updated = False
    #         already_marked = False
    #         student_info = None

    #         # Duyệt qua các dòng để tìm học sinh
    #         for row in sheet.iter_rows(min_row=2, max_col=column_letter):
    #             # So sánh tên (cột thứ 2, giả sử)
    #             cell_name = row[1].value if row[1] else None
                
    #             # So sánh tên không phân biệt chữ hoa chữ thường và loại bỏ khoảng trắng
    #             if cell_name and cell_name.strip().lower() == name.strip().lower():
    #                 # Kiểm tra xem đã điểm danh chưa
    #                 if row[column_letter-1].value:
    #                     already_marked = True
    #                     student_info = (row[0].value, cell_name)  # Mã số, tên
    #                     break
    #                 else:
    #                     # Điểm danh
    #                     sheet.cell(row=row[0].row, column=column_letter).value = f"Đã điểm danh - {datetime_string}"
    #                     updated = True
    #                     student_info = (row[0].value, cell_name)  # Mã số, tên
    #                     break

    #         # Xử lý kết quả điểm danh
    #         if already_marked:
    #             self.textEdit_missingStudents.setText(
    #                 f"Học sinh {student_info[1]} (Mã: {student_info[0]}) "
    #                 f"đã điểm danh vào ngày {self.selected_date}."
    #             )
    #         elif updated:
    #             # Thành công điểm danh
    #             self.textEdit_missingStudents.setText(
    #                 f"Điểm danh thành công:\n"
    #                 f"Mã: {student_info[0]}\n"
    #                 f"Tên: {student_info[1]}\n"
    #                 f"Ngày: {self.selected_date}\n"
    #                 f"Giờ: {datetime_string}"
    #             )
    #         else:
    #             # Không tìm thấy học sinh
    #             self.textEdit_missingStudents.setText(
    #                 f"Học sinh {name} không có trong danh sách lớp này."
    #             )

    #         # Lưu workbook
    #         workbook.save(self.excel_file)
    #         workbook.close()

    #     except Exception as e:
    #         print(f"Lỗi trong quá trình ghi vào Excel: {e}")

    def load_known_faces(self, directory):
        for filename in os.listdir(directory):
            if filename.endswith((".jpg", ".jpeg", ".png")):
                image_path = os.path.join(directory, filename)
                image = face_recognition.load_image_file(image_path)
                encoding = face_recognition.face_encodings(image)[0]
                name = os.path.splitext(filename)[0]
                
                self.encoded_known_faces.append(encoding)
                self.known_names.append(name)

    def start_camera(self):
        self.video_capture = cv2.VideoCapture(0)
        self.update_camera_frame()

    def update_camera_frame(self):
        ret, frame = self.video_capture.read()
        if not ret:
            messagebox.showerror("Lỗi", "Không thể đọc khung hình từ webcam")
            return

        # Resize frame for faster processing
        small_frame = cv2.resize(frame, (0, 0), fx=0.5, fy=0.5)
        rgb_small_frame = cv2.cvtColor(small_frame, cv2.COLOR_BGR2RGB)

        # Find faces
        current_face_locations = face_recognition.face_locations(rgb_small_frame)
        current_face_encodings = face_recognition.face_encodings(rgb_small_frame, current_face_locations)

        # Process each detected face
        for face_encoding, face_location in zip(current_face_encodings, current_face_locations):
            # Compare face with known faces
            matches = face_recognition.compare_faces(self.encoded_known_faces, face_encoding)
            face_distances = face_recognition.face_distance(self.encoded_known_faces, face_encoding)
            
            # Find best match
            best_match_index = np.argmin(face_distances)
            if face_distances[best_match_index] < 0.50:
                name = self.known_names[best_match_index].lower()
                self.mark_attendance(name)
            else:
                name = "Unknown"

            # Draw rectangle and name
            top, right, bottom, left = face_location
            top, right, bottom, left = top * 2, right * 2, bottom * 2, left * 2
            cv2.rectangle(frame, (left, top), (right, bottom), (0, 255, 0), 2)
            cv2.putText(frame, name, (left + 6, bottom - 6), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 2)

        # Convert to RGB for Tkinter
        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        photo = tk.PhotoImage(data=cv2.imencode('.png', frame_rgb)[1].tobytes())
        self.camera_label.configure(image=photo)
        self.camera_label.image = photo

        # Schedule next update
        self.window.after(10, self.update_camera_frame)

    def mark_attendance(self, name):
        if not self.excel_file:
            messagebox.showwarning("Lỗi", "Chưa chọn file Excel.")
            return
        
        if not self.selected_date:
            messagebox.showwarning("Lỗi", "Chưa chọn ngày điểm danh.")
            return

        try:
            workbook = openpyxl.load_workbook(self.excel_file)
            sheet = workbook.active
            column_letter = None

            # Tìm cột tương ứng với ngày
            for col in range(1, sheet.max_column + 1):
                cell_value = sheet.cell(row=1, column=col).value
                if cell_value == self.selected_date:
                    column_letter = col
                    break

            # Nếu không tìm thấy cột, tạo cột mới
            if column_letter is None:
                column_letter = sheet.max_column + 1
                sheet.cell(row=1, column=column_letter).value = self.selected_date

            # Ghi thời gian điểm danh
            now = datetime.now()
            datetime_string = now.strftime("%H:%M:%S")
            updated = False
            already_marked = False

            # Kiểm tra và cập nhật điểm danh
            for row in sheet.iter_rows(min_row=2, max_col=column_letter):
                cell_name = row[0].value
                if cell_name and cell_name.strip().lower() == name.strip().lower():
                    if row[column_letter-1].value:
                        already_marked = True
                        break
                    else:
                        sheet.cell(row=row[0].row, column=column_letter).value = f"Đã điểm danh - {datetime_string}"
                        updated = True
                        break

            # Cập nhật thông báo
            if already_marked:
                self.attendance_text.insert(tk.END, f"Học sinh {name} đã điểm danh vào ngày {self.selected_date}.\n")
            elif not updated:
                self.attendance_text.insert(tk.END, f"Học sinh {name} không có trong danh sách lớp này.\n")

            workbook.save(self.excel_file)
            workbook.close()

        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi trong quá trình ghi vào Excel: {e}")


    def toggle_camera(self):
        if hasattr(self, 'id_entry') and self.id_entry:
            camera_index = self.id_entry.get().strip()
            
            if not camera_index:
                tk.messagebox.showwarning("Lỗi", "Vui lòng nhập ID Camera.")
                return
                
            try:
                camera_index = int(camera_index)
            except ValueError:
                tk.messagebox.showwarning("Lỗi", "Giá trị ID không hợp lệ.")
                return
                
            if self.video_capture is None:
                self.video_capture = cv2.VideoCapture(camera_index)
                if not self.video_capture.isOpened():
                    tk.messagebox.showerror("Lỗi", "Không thể mở camera")
                    return
                self.camera_button.configure(text="Tắt Camera")
            else:
                self.video_capture.release()
                self.video_capture = None
                self.camera_canvas.delete("all")
                self.camera_button.configure(text="Bật Camera")
        else:
            tk.messagebox.showwarning("Lỗi", "Không tìm thấy đối tượng ID Camera.")

    def update_loop(self):
        """Main camera update loop"""
        if self.video_capture is not None and self.video_capture.isOpened():
            ret, frame = self.video_capture.read()
            if ret:
                try:
                    # Resize frame for faster processing
                    small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)
                    rgb_small_frame = cv2.cvtColor(small_frame, cv2.COLOR_BGR2RGB)
                    
                    # Process face recognition
                    face_locations = face_recognition.face_locations(rgb_small_frame)
                    if face_locations:
                        face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)
                        
                        for face_encoding, face_location in zip(face_encodings, face_locations):
                            if len(self.encoded_known_faces) > 0:
                                matches = face_recognition.compare_faces(self.encoded_known_faces, face_encoding, tolerance=0.6)
                                face_distances = face_recognition.face_distance(self.encoded_known_faces, face_encoding)
                                
                                if len(face_distances) > 0:
                                    best_match_index = np.argmin(face_distances)
                                    if matches[best_match_index]:
                                        name = self.known_names[best_match_index]
                                        
                                        # Check if enough time has passed since last attendance
                                        current_time = datetime.now()
                                        if (name not in self.last_attendance_time or 
                                            (current_time - self.last_attendance_time[name]).total_seconds() > self.attendance_cooldown):
                                            
                                            self.mark_attendance(name)
                                            self.last_attendance_time[name] = current_time
                                            # Print confirmation message
                                            print(f"Marked attendance for {name} at {current_time}")
                                    else:
                                        name = "Unknown"
                                else:
                                    name = "Unknown"
                                    
                                # Scale back face locations for display
                                top, right, bottom, left = face_location
                                top *= 4
                                right *= 4
                                bottom *= 4
                                left *= 4
                                
                                # Draw face box and name
                                color = (0, 255, 0) if name != "Unknown" else (0, 0, 255)
                                cv2.rectangle(frame, (left, top), (right, bottom), color, 2)
                                cv2.putText(frame, name, (left + 6, bottom - 6),
                                          cv2.FONT_HERSHEY_COMPLEX, 0.6, (255, 255, 255), 1)

                except Exception as e:
                    print(f"Error in face recognition: {e}")

                # Update display
                try:
                    frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    frame_pil = Image.fromarray(frame_rgb)
                    frame_tk = ImageTk.PhotoImage(image=frame_pil)
                    
                    self.camera_canvas.create_image(0, 0, anchor="nw", image=frame_tk)
                    self.camera_canvas.image = frame_tk
                except Exception as e:
                    print(f"Error updating display: {e}")

        self.window.after(20, self.update_loop)

    def select_excel_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            self.excel_file = filepath
            self.excel_entry.delete(0, "end")
            self.excel_entry.insert(0, os.path.basename(filepath))         
            
    def display_excel_data(self):
        if not self.excel_file:
            messagebox.showwarning("Lỗi", "Vui lòng chọn file Excel trước.")
            return

        try:
            self.dataframe = pd.read_excel(self.excel_file)

            # Clear existing items
            for item in self.tree.get_children():
                self.tree.delete(item)

            # Update columns for Treeview
            self.tree['columns'] = tuple(self.dataframe.columns)
            self.tree.column("#0", width=0, stretch=tk.NO)  # Hide the first column
            for column in self.dataframe.columns:
                self.tree.column(column, anchor=tk.W, width=150)
                self.tree.heading(column, text=column, anchor=tk.W)

            # Add rows
            for _, row in self.dataframe.iterrows():
                values = [row[col] for col in self.dataframe.columns]
                self.tree.insert("", tk.END, values=values)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi khi tải file Excel: {str(e)}")

    def mark_attendance(self, name):
        if self.dataframe is None:
            messagebox.showwarning("Lỗi", "Vui lòng chọn file Excel trước.")
            return

        try:
            selected_date = self.date_entry.get_date().strftime("%d/%m/%Y")

            # Thêm cột ngày nếu chưa tồn tại
            if selected_date not in self.dataframe.columns:
                self.dataframe[selected_date] = np.nan

            # Tìm và điểm danh
            name_column = self.dataframe.columns[1]  # Giả sử cột tên ở vị trí thứ 2
            mask = self.dataframe[name_column] == name
            
            if mask.any():
                self.dataframe.loc[mask, selected_date] = "Có mặt"
                
                # Lưu lại file Excel
                self.dataframe.to_excel(self.excel_file, index=False)
                messagebox.showinfo("Thành công", f"Đã điểm danh cho {name}.")
            else:
                messagebox.showwarning("Lỗi", f"Không tìm thấy học sinh tên {name}.")
        
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi khi điểm danh: {str(e)}")

    def show_unmarked(self):
        if self.dataframe is None:
            messagebox.showwarning("Lỗi", "Chưa tải dữ liệu từ file Excel.")
            return

        try:
            selected_date = self.date_entry.get_date().strftime("%d/%m/%Y")
            
            if selected_date not in self.dataframe.columns:
                messagebox.showerror("Lỗi", "Ngày được chọn chưa tồn tại trong dữ liệu.")
                return

            unmarked_students = self.dataframe[pd.isna(self.dataframe[selected_date])]
            self.update_treeview(unmarked_students)
        
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))

    def show_marked(self):
        if self.dataframe is None:
            messagebox.showwarning("Lỗi", "Chưa tải dữ liệu từ file Excel.")
            return

        try:
            selected_date = self.date_entry.get_date().strftime("%d/%m/%Y")
            
            if selected_date not in self.dataframe.columns:
                messagebox.showerror("Lỗi", "Ngày được chọn chưa tồn tại trong dữ liệu.")
                return

            marked_students = self.dataframe[self.dataframe[selected_date] == "Có mặt"]
            self.update_treeview(marked_students)
        
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))

    def send_email_to_parents(self):
        if self.dataframe is None:
            messagebox.showwarning("Lỗi", "Chưa tải dữ liệu từ file Excel.")
            return

        try:
            selected_date = self.date_entry.get_date().strftime("%d/%m/%Y")
            
            # Kiểm tra cột email
            if len(self.dataframe.columns) < 5:
                messagebox.showerror("Lỗi", "File Excel không đủ cột.")
                return

            email_column = self.dataframe.columns[4]  # Cột email
            name_column = self.dataframe.columns[1]   # Cột tên

            # Lọc các học sinh vắng mặt
            unmarked_students = self.dataframe[pd.isna(self.dataframe[selected_date])]

            if unmarked_students.empty:
                messagebox.showinfo("Thông báo", "Không có học sinh vắng mặt.")
                return

            # Cấu hình email
            sender_email = "vobaolong15@gmail.com"
            sender_password = "wpmrfsgbhsfgwiin"

            # Gửi email cho từng phụ huynh
            for _, student in unmarked_students.iterrows():
                parent_email = student[email_column]
                student_name = student[name_column]

                if pd.notna(parent_email):
                    try:
                        # Tạo email
                        msg = MIMEMultipart()
                        msg['From'] = sender_email
                        msg['To'] = parent_email
                        msg['Subject'] = f"Thông báo vắng học của {student_name}"

                        body = f"""
                        Kính gửi Quý Phụ huynh,

                        Nhà trường xin thông báo: Em {student_name} đã vắng mặt vào ngày {selected_date}.

                        Trân trọng,
                Trường THPT Nguyễn Thượng Hiền
                        """

                        msg.attach(MIMEText(body, 'plain', 'utf-8'))

                        # Kết nối và gửi email
                        with smtplib.SMTP('smtp.gmail.com', 587) as server:
                            server.starttls()
                            server.login(sender_email, sender_password)
                            server.send_message(msg)

                        print(f"Đã gửi email tới {parent_email}")

                    except Exception as e:
                        print(f"Lỗi gửi email tới {parent_email}: {e}")

            messagebox.showinfo("Thành công", "Đã gửi email thông báo.")

        except Exception as e:
            messagebox.showerror("Lỗi", str(e))

    def update_treeview(self, filtered_dataframe):
        # Xóa các item hiện tại
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Thêm dữ liệu đã lọc
        for _, row in filtered_dataframe.iterrows():
            values = [str(row[col]) for col in self.dataframe.columns]
            self.tree.insert("", tk.END, values=values)

    def clear_log(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
def main():
    app = CameraApp()
    app.window.mainloop()

if __name__ == "__main__":
    main()