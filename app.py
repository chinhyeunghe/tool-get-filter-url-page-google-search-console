import customtkinter as ctk
import os
import sys
from PIL import Image
from tkinter import filedialog
from tkcalendar import DateEntry
import tkinter.font as tkFont
import threading
import time
import datetime

# Import thư viện Google API
from googleapiclient.discovery import build
# Import thư viện xác thực tài khoản dịch vụ
from oauth2client.service_account import ServiceAccountCredentials
# Import thư viện xử lý dữ liệu dạng bảng
import pandas as pd

ctk.set_appearance_mode("dark")  # Giao diện dark mode
ctk.set_default_color_theme("blue")

def resource_path(relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Dang Duc Chinh (DevC) - Tool siêu nhân lấy URL từ Google Search Console")
        self.geometry("800x700")
        self.grid_columnconfigure(0, weight=1)

        # Load ảnh
        self.bg_image = ctk.CTkImage(Image.open(resource_path("images-FS.png")), size=(80, 80))
        self.bg_label = ctk.CTkLabel(self, image=self.bg_image, text="")
        self.bg_label.place(x=15, y=15)

        # current time
        current_year = datetime.datetime.now().year
        before_year = current_year - 1

        # thanh tiến trình
        self.progressbar = ctk.CTkProgressBar(self, orientation="horizontal", width=400)
        self.progressbar.grid(row=16, column=0, padx=20, pady=10)
        self.progressbar.set(0)  # khởi tạo ở 0
        self.progressbar.grid_remove()  # ẩn thanh lúc đầu



        # Label & input URL
        self.lable_cua_toi = ctk.CTkLabel(self, text="Nhập URL website của bạn.", text_color="white")
        self.lable_cua_toi.grid(row=0, column=0, padx=20, pady=(40, 10))

        self.url_entry = ctk.CTkEntry(self, width=400, height=30, placeholder_text="URL website. Ví dụ https://example.com",
                                      border_color="#FFFFFF", border_width=1,  text_color="white")
        self.url_entry.grid(row=1, column=0, padx=20, pady=10)

        # Chọn file JSON
        self.my_file_json = ctk.CTkButton(self, text="Chọn file JSON xác thực", width=400, command=self.handle_choose_file,
                                           border_color="white", border_width=1, text_color="white", fg_color="#57C785")
        self.my_file_json.grid(row=3, column=0, padx=20, pady=10)

        # Chọn ngày bắt đầu và kết thúc
        font_style = tkFont.Font(family="Arial", size=16)

        self.start_label = ctk.CTkLabel(self, text="Ngày bắt đầu:", text_color="white")
        self.start_label.grid(row=4, column=0)
        self.start_date = DateEntry(self, width=15, background='darkblue', foreground='white',
                                    borderwidth=2, year=before_year, date_pattern='yyyy-mm-dd', font=font_style)
        self.start_date.grid(row=5, column=0)

        self.end_label = ctk.CTkLabel(self, text="Ngày kết thúc:", text_color="white")
        self.end_label.grid(row=6, column=0)
        self.end_date = DateEntry(self, width=15, background='darkblue', foreground='white',
                                  borderwidth=2, year=current_year, date_pattern='yyyy-mm-dd', font=font_style)
        self.end_date.grid(row=7, column=0)

        # Lọc theo từ khóa
        self.lable_filter = ctk.CTkLabel(self, text="Lọc theo từ khóa (Mặc định lấy tất cả nếu không điền)",
                                         bg_color="transparent", text_color="white")
        self.lable_filter.grid(row=8, column=0, padx=10, pady=(40, 10))

        self.kw_filter = ctk.CTkEntry(self, width=400, height=30, placeholder_text="Mỗi từ khóa cách nhau bởi dấu phẩy",
                                      border_color="white", border_width=1,  text_color="white")
        self.kw_filter.grid(row=9, column=0, padx=20, pady=10)

        # Xác định kiểu validate: chỉ cho phép số
        vcmd = (self.register(self.validate_number), '%P')

        # Label chung cho phân trang
        self.pagination_label = ctk.CTkLabel(self, text="Phân trang (Vị trí bắt đầu - Số lượng dòng). Mặc định 0 - 25.000 và không nên lấy vượt quá 50.000", text_color="white")
        self.pagination_label.grid(row=10, column=0, padx=20, pady=(40, 10))

        # Tạo frame để chứa hai input trên cùng một dòng
        self.pagination_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.pagination_frame.grid(row=11, column=0, padx=20, pady=10)

        # Input: Vị trí bắt đầu
        self.start_from_input = ctk.CTkEntry(self.pagination_frame, validate="key", validatecommand=vcmd, width=180, height=30, placeholder_text="Vị trí bắt đầu",
                                            border_color="#FFFFFF", border_width=1, text_color="white")
        self.start_from_input.grid(row=0, column=0, padx=(0, 10), pady=10)

        # Input: Số lượng dòng
        self.num_page_input = ctk.CTkEntry(self.pagination_frame, validate="key", validatecommand=vcmd, width=180, height=30, placeholder_text="Số dòng muốn lấy",
                                        border_color="#FFFFFF", border_width=1, text_color="white")
        self.num_page_input.grid(row=0, column=1, padx=(10, 0), pady=10)

        # Nút xử lý
        self.process_button = ctk.CTkButton(self, text="Bắt đầu xử lý", command=self.start_processing, width=300, height=35,
                                             border_color="white", border_width=1, text_color="white")
        self.process_button.grid(row=15, column=0, padx=20, pady=20)
    
    def validate_number(self, value):
        if value == "" or value.isdigit():
            return True
        else:
            return False


    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def handle_choose_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if file_path:
            self.json_file_path = file_path
            print("Bạn đã chọn file:", file_path)

    def start_processing(self):
        if not self.json_file_path:
            print("Vui lòng chọn file JSON trước!")
            return

        if not self.url_entry.get():
            print("Vui lòng nhập URL website!")
            return

        self.process_button.configure(text="Đang xử lý...", state="disabled")
        threading.Thread(target=self.process_api_logic).start()

    def start_processing(self):
        if not hasattr(self, "json_file_path"):
            print("Vui lòng chọn file JSON trước!")
            return

        if not self.url_entry.get():
            print("Vui lòng nhập URL website!")
            return

        self.process_button.configure(text="Đang xử lý...", state="disabled")

        # Hiện progress bar
        self.progressbar.grid()
        self.progressbar.set(0)
        self.after(100, self.run_progress)  # bắt đầu chạy progress bar

        threading.Thread(target=self.process_api_logic).start()

    def run_progress(self):
        current_value = self.progressbar.get()
        new_value = current_value + 0.01
        if new_value > 1.0:
            new_value = 0
        self.progressbar.set(new_value)
        
        # Dùng cget để lấy trạng thái thay vì truy cập trực tiếp
        if self.process_button.cget("state") == "disabled":
            self.after(50, self.run_progress)



    def process_api_logic(self):
        try:
            print("Bắt đầu xử lý API...")

            url_site = self.url_entry.get()
            start_date = self.start_date.get_date().strftime('%Y-%m-%d')
            end_date = self.end_date.get_date().strftime('%Y-%m-%d')
            keyword_filter = self.kw_filter.get().strip()

            start_row = int(self.start_from_input.get() or 0)
            row_limit = int(self.num_page_input.get() or 2500)
            row_limit = min(row_limit, 25000)  # giới hạn API

            SCOPES = ['https://www.googleapis.com/auth/webmasters.readonly']
            credentials = ServiceAccountCredentials.from_json_keyfile_name(self.json_file_path, scopes=SCOPES)
            service = build('searchconsole', 'v1', credentials=credentials)

            all_rows = []

            max_api_calls = 20  # giới hạn số lần gọi API để tránh gọi quá nhiều

            def fetch_data_for_keyword(kw, start_row, max_rows):
                current_start_row = start_row
                api_call_count = 0
                collected_rows = []
                while True:
                    if api_call_count >= max_api_calls:
                        print(f"Đã đạt giới hạn số lần gọi API tối đa {max_api_calls}, dừng từ khóa '{kw}'.")
                        break

                    rows_to_fetch = min(row_limit, max_rows - len(collected_rows))

                    request = {
                        'startDate': start_date,
                        'endDate': end_date,
                        'dimensions': ['query', 'page'],
                        'rowLimit': rows_to_fetch,
                        'startRow': current_start_row,
                        'dimensionFilterGroups': [{
                            "groupType": "and",
                            "filters": [{
                                "dimension": "query",
                                "operator": "contains",  # có thể đổi "equals" nếu cần chính xác
                                "expression": kw
                            }]
                        }]
                    }

                    print(f"Lấy dữ liệu từ {current_start_row} với {rows_to_fetch} dòng cho từ khóa '{kw}'")
                    response = service.searchanalytics().query(siteUrl=url_site, body=request).execute()

                    if 'rows' not in response:
                        print(f"Không có dữ liệu trả về cho từ khóa '{kw}'.")
                        break

                    rows = response['rows']
                    collected_rows.extend(rows)

                    if len(rows) < rows_to_fetch:
                        print(f"Đã lấy hết dữ liệu cho từ khóa '{kw}'.")
                        break

                    if len(collected_rows) >= max_rows:
                        print(f"Đã lấy đủ {max_rows} dòng cho từ khóa '{kw}'.")
                        break

                    current_start_row += rows_to_fetch
                    api_call_count += 1

                    time.sleep(0.5)
                return collected_rows

            if keyword_filter:
                keywords = [kw.strip() for kw in keyword_filter.split(',') if kw.strip()]
                for kw in keywords:
                    rows_for_kw = fetch_data_for_keyword(kw, start_row, row_limit)
                    all_rows.extend(rows_for_kw)
            else:
                current_start_row = start_row
                api_call_count = 0
                while True:
                    if api_call_count >= max_api_calls:
                        print(f"Đã đạt giới hạn số lần gọi API tối đa {max_api_calls}, dừng.")
                        break

                    rows_to_fetch = min(row_limit, row_limit - len(all_rows))

                    request = {
                        'startDate': start_date,
                        'endDate': end_date,
                        'dimensions': ['query', 'page'],
                        'rowLimit': rows_to_fetch,
                        'startRow': current_start_row
                    }

                    print(f"Lấy dữ liệu từ {current_start_row} với {rows_to_fetch} dòng")
                    response = service.searchanalytics().query(siteUrl=url_site, body=request).execute()

                    if 'rows' not in response:
                        print("Không có dữ liệu trả về.")
                        break

                    rows = response['rows']
                    all_rows.extend(rows)

                    if len(rows) < rows_to_fetch:
                        print("Đã lấy hết dữ liệu.")
                        break

                    if len(all_rows) >= row_limit:
                        print(f"Đã lấy đủ {row_limit} dòng, dừng.")
                        break

                    current_start_row += rows_to_fetch
                    api_call_count += 1

                    time.sleep(0.5)

            # Xử lý dữ liệu vào DataFrame
            data = []
            for row in all_rows[:row_limit]:  # đảm bảo chỉ lấy tối đa row_limit dòng
                keyword = row['keys'][0]
                url = row['keys'][1]
                clicks = row.get('clicks', 0)
                impressions = row.get('impressions', 0)
                ctr = row.get('ctr', 0)
                position = row.get('position', 0)
                data.append({
                    'Keyword': keyword,
                    'URL': url,
                    'Clicks': clicks,
                    'Impressions': impressions,
                    'CTR': ctr,
                    'Position': position
                })

            df = pd.DataFrame(data)

            now = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
            file_name = f"Output_file_{now}.csv"

            # Tạo folder 'output' nếu chưa tồn tại
            output_folder = "outputs"
            os.makedirs(output_folder, exist_ok=True)

            # Đặt tên file với đường dẫn folder 'output'
            file_name = os.path.join(output_folder, f"Output_file_{now}.csv")
            df.to_csv(file_name, index=False, encoding='utf-8-sig')

            print(f"Đã xuất dữ liệu thành công vào file {file_name}")

        except Exception as e:
            print("Đã xảy ra lỗi:", e)

        finally:
            self.process_button.configure(text="Bắt đầu xử lý", state="normal")
            self.progressbar.grid_remove()



app = App()
app.mainloop()
