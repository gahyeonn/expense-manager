import tkinter as tk
from tkinter import filedialog
from parse import ExcelProcessor as processor

class ERPSeparator:
    def __init__(self, root):
        self.root = root
        self.file_path = None
        self.saving_path = None
        
        self.root.title("ERP 전표 분리기")
        
        self.reading_button = tk.Button(root, text="비용 처리할 엑셀 파일 선택", command=self.select_excel_file)
        self.reading_button.pack(pady=15)
        
        self.reading_label = tk.Label(root, text="")
        self.reading_label.pack(pady=5)
        
        
        self.saving_button = tk.Button(root, text="엑셀 파일 저장할 폴더 선택", command=self.select_saving_directory)
        self.saving_button.pack(pady=15)
        
        self.saving_label = tk.Label(root, text="")
        self.saving_label.pack(pady=5)
        
        self.submit_button = tk.Button(root, text="완료", command=self.start_parser)
        self.submit_button.pack(pady=20)
        
    # 비용 처리를 위한 엑셀 파일 선택
    def select_excel_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.file_path:
            self.reading_label.config(text=f"선택한 파일: {self.file_path}")
    
    # 계정코드 별 분리된 엑셀 파일 저장 경로 설정
    def select_saving_directory(self):
        self.saving_path = filedialog.askdirectory()
        if self.saving_path:
            self.saving_label.config(text=f"선택한 폴더: {self.saving_path}")
    
    def start_parser(self): #임시 코드
        passer = processor()
        processor.process_excel(passer, self.file_path, self.saving_path)


if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("400x300")
    app = ERPSeparator(root)
    root.mainloop() #GUI 실행