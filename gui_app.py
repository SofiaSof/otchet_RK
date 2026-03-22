import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from download_photos import main as download_main


class PhotoDownloaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Загрузчик фото - Russ Outdoor")
        self.root.geometry("600x400")
        self.root.resizable(False, False)
        
        self.is_running = False
        self.log_lines = []
        self.max_log_lines = 100
        
        self.setup_ui()
    
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Загрузчик фотографий с Russ Outdoor", 
                 font=("Arial", 14, "bold")).pack(pady=(0, 20))
        
        config_frame = ttk.LabelFrame(main_frame, text="Настройки", padding="10")
        config_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(config_frame, text="Диапазон строк (через дефис):").grid(
            row=0, column=0, sticky=tk.W, pady=5)
        self.rows_entry = ttk.Entry(config_frame, width=15)
        self.rows_entry.insert(0, "8-103")
        self.rows_entry.grid(row=0, column=1, sticky=tk.W, padx=10, pady=5)
        
        ttk.Label(config_frame, text="Путь к Excel:").grid(
            row=1, column=0, sticky=tk.W, pady=5)
        self.excel_path = ttk.Entry(config_frame, width=40)
        self.excel_path.insert(0, "C:/test/Otchet_Samocat.xlsx")
        self.excel_path.grid(row=1, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Button(config_frame, text="...", width=4,
                   command=self.browse_excel).grid(row=1, column=2, pady=5)
        
        ttk.Label(config_frame, text="Папка для фото:").grid(
            row=2, column=0, sticky=tk.W, pady=5)
        self.photos_path = ttk.Entry(config_frame, width=40)
        self.photos_path.insert(0, "C:/test/photos")
        self.photos_path.grid(row=2, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Button(config_frame, text="...", width=4,
                   command=self.browse_photos).grid(row=2, column=2, pady=5)
        
        self.btn_frame = ttk.Frame(main_frame)
        self.btn_frame.pack(pady=10)
        
        self.start_btn = ttk.Button(self.btn_frame, text="Начать загрузку", 
                                     command=self.start_download)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_btn = ttk.Button(self.btn_frame, text="Стоп", 
                                   command=self.stop_download, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.pack(fill=tk.X, pady=(10, 5))
        
        log_frame = ttk.LabelFrame(main_frame, text="Лог", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD,
                                font=("Courier", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
    
    def browse_excel(self):
        path = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self.excel_path.delete(0, tk.END)
            self.excel_path.insert(0, path)
    
    def browse_photos(self):
        path = filedialog.askdirectory(title="Выберите папку для фото")
        if path:
            self.photos_path.delete(0, tk.END)
            self.photos_path.insert(0, path)
    
    def log(self, message):
        self.log_lines.append(message)
        if len(self.log_lines) > self.max_log_lines:
            self.log_lines.pop(0)
        self.log_text.delete(1.0, tk.END)
        self.log_text.insert(tk.END, "\n".join(self.log_lines))
        self.log_text.see(tk.END)
    
    def start_download(self):
        rows_range = self.rows_entry.get().strip()
        if not rows_range:
            messagebox.showerror("Ошибка", "Укажите диапазон строк")
            return
        
        try:
            if "-" in rows_range:
                start, end = rows_range.split("-")
                self.row_nums = list(range(int(start.strip()), int(end.strip()) + 1))
            else:
                self.row_nums = [int(rows_range.strip())]
        except ValueError:
            messagebox.showerror("Ошибка", "Неверный формат диапазона строк")
            return
        
        self.is_running = True
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.progress.start()
        
        self.log(f"Начало загрузки: строки {self.row_nums[0]}-{self.row_nums[-1]}")
        
        self.thread = threading.Thread(target=self.run_download, daemon=True)
        self.thread.start()
    
    def run_download(self):
        try:
            original_stdout = sys.stdout
            sys.stdout = self
            
            download_main(
                excel_path=self.excel_path.get(),
                photos_path=self.photos_path.get(),
                row_nums=self.row_nums
            )
            
            self.log("Загрузка завершена!")
            self.root.after(0, lambda: messagebox.showinfo("Готово", "Загрузка фотографий завершена!"))
            
        except Exception as e:
            self.log(f"Ошибка: {e}")
            self.root.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
        finally:
            self.is_running = False
            self.progress.stop()
            self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))
            sys.stdout = original_stdout
    
    def write(self, message):
        self.root.after(0, self.log, message.strip())
    
    def stop_download(self):
        self.is_running = False
        self.log("Остановка...")
        messagebox.showinfo("Информация", "Нажмите Ctrl+C в консоли для остановки браузера")


def main():
    root = tk.Tk()
    app = PhotoDownloaderGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
