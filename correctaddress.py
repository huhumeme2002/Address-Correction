import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import google.generativeai as genai
import time
import threading
import re

def start_processing(self):
    threading.Thread(target=self.process_addresses, daemon=True).start()

# Danh sách API key
API_KEYS = [
    '',  
    '',
    ''   
]

def setup_model(api_key):
    genai.configure(api_key=api_key)
    return genai.GenerativeModel('gemini-2.0-flash')

DEFAULT_PROMPT = """Hãy kiểm tra các địa chỉ sau:
- Nếu địa chỉ thuộc thành phố Hồ Chí Minh, hãy sửa chính tả và định dạng lại địa chỉ theo chuẩn Việt Nam và thêm "Hồ Chí Minh" ở cuối.
- Nếu địa chỉ không thuộc thành phố Hồ Chí Minh, hãy tự xác định tên tỉnh của địa chỉ đó, sửa chính tả và định dạng lại, và đặt tên tỉnh tìm được ở cuối.
Chỉ trả về kết quả duy nhất là địa chỉ đã sửa, không có dòng nào khác, không cần đánh số thứ tự câu trả lời.
Địa chỉ đã sửa (có dạng "[Số nhà] [Tên đường], [Phường/Xã], [Quận/Huyện], [Tên tỉnh/thành phố]")
Ví dụ:
Input:
362/25/30F Phan Huy Ích, Phường 12, Quận Gò Vấp, TP. HCM
Output:
362/25/30F Phan Huy Ích, Phường 12, Quận Gò Vấp, Hồ Chí Minh
Danh sách địa chỉ:"""

class AddressCorrectorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Address Corrector")
        self.root.geometry("600x500")
        self.file_path = None

        # Tạo danh sách quản lý API key: mỗi key có thông tin counter và next_available.
        self.keys_info = [
            {'api_key': key, 'counter': 0, 'next_available': 0}
            for key in API_KEYS
        ]
        self.current_api_key = self.keys_info[0]['api_key']
        self.model = setup_model(self.current_api_key)
        
        # Số địa chỉ xử lý mỗi lần gọi API (batch size)
        self.batch_size = 40  
        self.setup_ui()
        
    def setup_ui(self):
        ttk.Label(self.root, text="Base Prompt:").pack(pady=5)
        self.prompt_text = tk.Text(self.root, height=10, width=80)
        self.prompt_text.insert(tk.END, DEFAULT_PROMPT)
        self.prompt_text.pack(pady=5)
        
        ttk.Button(self.root, text="Chọn file Excel", command=self.load_file).pack(pady=10)
        # Dùng start_processing để chạy xử lý trên luồng nền
        ttk.Button(self.root, text="Xử lý địa chỉ", command=self.start_processing).pack(pady=10)
        self.status_label = ttk.Label(self.root, text="")
        self.status_label.pack(pady=5)
        
        # Label hiển thị trạng thái của các API key
        self.api_status_label = ttk.Label(self.root, text="", justify="left", font=("Consolas", 10))
        self.api_status_label.pack(pady=5)
        
    def safe_update_status(self, text):
        self.root.after(0, lambda: self.status_label.config(text=text))
        
    def update_api_status(self):
        now = time.time()
        lines = []
        for idx, key_info in enumerate(self.keys_info):
            key_display = key_info['api_key'][-6:]
            counter = key_info['counter']
            if now < key_info['next_available']:
                cooldown = int(key_info['next_available'] - now)
                status = f"Cooldown ({cooldown}s)"
            else:
                status = "Available"
            lines.append(f"API Key {idx+1} ({key_display}): {status}, counter: {counter}")
        text = "\n".join(lines)
        self.root.after(0, lambda: self.api_status_label.config(text=text))
    
    def load_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.file_path:
            self.safe_update_status(f"Đã chọn file: {self.file_path}")
    
    def get_available_key(self):
        while True:
            now = time.time()
            for key_info in self.keys_info:
                if now >= key_info['next_available']:
                    return key_info
            next_ready = min(key_info['next_available'] for key_info in self.keys_info)
            wait_time = max(0, next_ready - now)
            self.safe_update_status(f"Tất cả API key đang nghỉ, chờ {int(wait_time)} giây...")
            time.sleep(wait_time)
    
    def start_processing(self):
        threading.Thread(target=self.process_addresses, daemon=True).start()
    
    def process_addresses(self):
        if not self.file_path:
            messagebox.showerror("Lỗi", "Vui lòng chọn file Excel trước.")
            return
        
        try:
            # Đọc file Excel và chỉ lấy cột có header "địa chỉ" (không phân biệt chữ hoa chữ thường)
            df = pd.read_excel(self.file_path)
            address_column = None
            for col in df.columns:
                if str(col).strip().lower() == "địa chỉ":
                    address_column = col
                    break
            if address_column is None:
                messagebox.showerror("Lỗi", "File Excel không có cột 'địa chỉ'.")
                return
            
            corrected_results = []
            province_results = []
            addresses = df[address_column].tolist()
            base_prompt = self.prompt_text.get("1.0", tk.END).strip()
            total = len(addresses)
            current_index = 0
            
            while current_index < total:
                batch_addresses = addresses[current_index: current_index + self.batch_size]
                
                # Xây dựng prompt cho batch này
                prompt_lines = [base_prompt]
                for i, addr in enumerate(batch_addresses, start=1):
                    cleaned_addr = str(addr).strip()
                    prompt_lines.append(f"{i}. {cleaned_addr}")
                prompt_lines.append("Output:")
                prompt = "\n".join(prompt_lines)
                
                # Lấy API key sẵn sàng (nếu key hiện tại đang cooldown, chọn key khác)
                key_info = self.get_available_key()
                if self.current_api_key != key_info['api_key']:
                    self.current_api_key = key_info['api_key']
                    self.model = setup_model(self.current_api_key)
                    self.safe_update_status(f"Sử dụng API key: {self.current_api_key}")
                
                # Gọi API và xử lý kết quả
                response = self.model.generate_content(prompt)
                output_text = response.text.strip()
                # Giả sử mỗi địa chỉ sẽ trả về 1 dòng kết quả
                output_lines = [line.strip() for line in output_text.splitlines() if line.strip()]
                expected_lines = len(batch_addresses)
                if len(output_lines) < expected_lines:
                    output_lines += [""] * (expected_lines - len(output_lines))
                elif len(output_lines) > expected_lines:
                    output_lines = output_lines[:expected_lines]
                
                # Với mỗi kết quả, trích xuất tên tỉnh từ dấu phẩy cuối cùng
                for line in output_lines:
                    corrected = re.sub(r'^\d+\.\s*', '', line)
                    # Lấy phần sau dấu phẩy cuối cùng
                    if "," in corrected:
                        province = corrected.rsplit(",", 1)[-1].strip()
                    else:
                        province = "Không xác định"
                    corrected_results.append(corrected)
                    province_results.append(province)
                
                current_index += len(batch_addresses)
                
                # Cập nhật counter cho API key được sử dụng
                key_info['counter'] += 1
                if key_info['counter'] >= 14:
                    key_info['next_available'] = time.time() + 60
                    key_info['counter'] = 0
                    self.safe_update_status(f"API key {key_info['api_key']} nghỉ 60 giây.")
                    # Không block toàn bộ luồng, chỉ đặt cooldown
                self.safe_update_status(f"Đã xử lý {current_index}/{total} địa chỉ.")
                self.update_api_status()
            
            # Lưu kết quả vào file Excel (thêm 2 cột: Địa chỉ đã sửa, Tên tỉnh)
            df['Địa chỉ đã sửa'] = corrected_results
            df['Tên tỉnh'] = province_results
            with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='a') as writer:
                df.to_excel(writer, sheet_name='Corrected_Addresses', index=False)
            self.safe_update_status("Xử lý hoàn tất!")
            messagebox.showinfo("Thành công", "Xử lý hoàn tất!\nKết quả được lưu trong sheet 'Corrected_Addresses'.")
            
        except Exception as e:
            try:
                df['Địa chỉ đã sửa'] = corrected_results + [""] * (len(df) - len(corrected_results))
                df['Tên tỉnh'] = province_results + [""] * (len(df) - len(province_results))
                with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='a') as writer:
                    df.to_excel(writer, sheet_name='Partial_Output', index=False)
                messagebox.showerror("Lỗi", f"Đã xảy ra lỗi: {str(e)}\nOutput tạm thời được lưu tại sheet 'Partial_Output'.")
            except Exception as save_error:
                messagebox.showerror("Lỗi", f"Lỗi: {str(e)}\nKhông thể lưu output tạm thời do: {str(save_error)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AddressCorrectorApp(root)
    root.mainloop()
