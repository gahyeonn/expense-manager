import tkinter as tk
from tkinter import filedialog
from parse import ExcelProcessor as processor

class ERPSeparator:
    def __init__(self, root):
        self.root = root
        self.file_path = None
        self.saving_path = None
        
        self.root.title("계정 과목별 엑셀 파일 분리기")
        
        self.blank_left = tk.Label(root, text="", width=5,  height=4)
        self.blank_left.grid(row=0, column=0)
        
        # 비용 처리할 엑셀 파일 선택
        self.reading_text = tk.Label(root, text="비용 처리할 엑셀 파일 선택: ")
        self.reading_text.grid(row=0, column=1, padx=0, pady=10)
        
        self.reading_label = tk.Label(root, height=1, width=50, relief="groove", background="white")
        self.reading_label.grid(row=0, column=2, padx=10, pady=10, ipady=3)

        self.reading_button = tk.Button(root, text="선택", width=10, command=self.select_excel_file)
        self.reading_button.grid(row=0, column=3)
        
        self.blank_right = tk.Label(root, text="", width=5)
        self.blank_right.grid(row=0, column=4)
        
        # 파일 저장할 폴더 선택
        self.saving_text = tk.Label(root, text="엑셀 파일 저장할 폴더 선택: ")
        self.saving_text.grid(row=1, column=1, padx=0, pady=10)
        
        self.saving_label = tk.Label(root, height=1, width= 50, relief="groove", background="white")
        self.saving_label.grid(row=1, column=2, padx=10, pady=10, ipady=3)
        
        self.saving_button = tk.Button(root, text="선택", width=10, command=self.select_saving_directory)
        self.saving_button.grid(row=1, column=3)
        
        # 엑셀 파일 분리 시작 버튼
        self.start_button = tk.Button(root, text="엑셀 분리 시작", width=15, command=self.start_parser)
        self.start_button.grid(row=2, column=2, pady=20)
        
        self.result_label = tk.Label(root, text="")
        self.result_label.grid(row=3, column=2, pady=10)
        
    # 비용 처리를 위한 엑셀 파일 선택
    def select_excel_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")], title="엑셀 파일 선택")
        if self.file_path:
            self.reading_label.config(text=self.file_path, anchor="w")
    
    # 계정코드 별 분리된 엑셀 파일 저장 경로 설정
    def select_saving_directory(self):
        self.saving_path = filedialog.askdirectory(title="엑셀 파일 저장할 폴더 선택")
        if self.saving_path:
            self.saving_label.config(text=self.saving_path, anchor="w")
    
    # 엑셀 파일 처리 시작
    def start_parser(self):
        passer = processor(self.file_path, self.saving_path)
        result = passer.process_excel()
        self.result_label.config(text=result, justify="left")


if __name__ == "__main__":
    root = tk.Tk()
    app = ERPSeparator(root)
    root.mainloop() #GUI 실행