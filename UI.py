import tkinter as tk
from tkinter import filedialog

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()
        self.create_console()
        self.output_directory = None
        self.sud = None
        self.range1 = None
        self.range2 = None

    def create_widgets(self):
        self.select_file = tk.Button(self)
        self.select_file["text"] = "Excel 파일 선택"
        self.select_file["command"] = self.select_excel_file
        self.select_file.pack(side="top")
        
        self.select_output_dir = tk.Button(self)  # 추가: 저장될 경로 선택 버튼
        self.select_output_dir["text"] = "저장 경로 선택"
        self.select_output_dir["command"] = self.select_output_directory
        self.select_output_dir.pack(side="top")
    
        self.select_sud = tk.Button(self)
        self.select_sud["text"] = "소리 설정"
        self.select_sud["command"] = self.select_sound_ratio
        self.select_sud.pack(side="top")
        
        self.select_range1 = tk.Button(self)
        self.select_range1["text"] = "최소 범위"
        self.select_range1["command"] = self.select_sound_range1
        self.select_range1.pack(side="top")

        self.select_range2 = tk.Button(self)
        self.select_range2["text"] = "최대 범위"
        self.select_range2["command"] = self.select_sound_range2
        self.select_range2.pack(side="top")

        self.start_process = tk.Button(self)
        self.start_process["text"] = "시작"
        self.start_process["command"] = self.start_process_func
        self.start_process.pack(side="top")

        self.quit = tk.Button(self, text="QUIT", fg="red",
                              command=self.master.destroy)
        self.quit.pack(side="bottom")

    def create_console(self):
        self.console = tk.Text(self)
        self.console.pack(side="top", fill="both", expand=True)

    def log(self, message):
        self.console.insert("end", message)
        self.console.see("end")

    def get_values(self):
        return {
            "efn": getattr(self, "efn", None),
            "output_directory": self.output_directory,
            "sud": self.sud,
            "range1": self.range1,
            "range2": self.range2
        }

    def select_excel_file(self):
        self.efn = filedialog.askopenfilename(initialdir="/", title="Select file",
                                              filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
        self.log(f"Selected Excel File: {self.efn}\n")
        
    def select_output_directory(self):
        self.output_directory = filedialog.askdirectory(title="Select directory to save processed videos")
        if not self.output_directory:
            self.log("No output directory selected.\n")
        else:
            self.log(f"Selected Output Directory: {self.output_directory}\n")

    def select_sound_ratio(self):
        self.sud = tk.simpledialog.askstring("Input", "소리 비율 조절 (2~2.5 추천, 소숫점 가능):",
                                             parent=self.master)
        self.log(f"Selected Sound Ratio: {self.sud}\n")

    def select_sound_range1(self):
        self.range1 = tk.simpledialog.askstring("Input","최소 1이상",parent=self.master)
        self.log(f"Selected 최소 범위: {self.range1}\n")

    def select_sound_range2(self):
        self.range2 = tk.simpledialog.askstring("Input","최소값 보단 높게 할 것, 최대 범위를 원하면 (M/m)을 입력",parent=self.master)
        self.log(f"Selected 최대 범위: {self.range2}\n")