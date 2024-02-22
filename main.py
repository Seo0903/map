import UI
from youtubeDLP import SoundDownloader
import soundController
import tkinter as tk

if __name__ == "__main__":   
    try:
        import sys
        import subprocess
        import tkinter
        import os
        import pydub
        import openpyxl
        import pandas
        import yt_dlp
    except Exception:
        # pip 모듈 업그레이드
        subprocess.check_call([sys.executable,'-m', 'pip', 'install', '--upgrade', 'pip'], shell=True)
        # 에러 발생한 모듈 설치
        subprocess.check_call([sys.executable,'-m', 'pip', 'install', '--upgrade', 'pandas'], shell=True)
        # 다시 import
        import sys
        import subprocess
        import tkinter
        import os
        import pydub
        import openpyxl
        import pandas
        import yt_dlp


    root = tk.Tk()
    app = UI.Application(master=root)
    app.mainloop()

    ui_values = app.get_values()

    # SoundController 클래스의 인스턴스 생성
    sound_downloader = SoundDownloader()
    sound_downloader.process_values(ui_values)
    sound_downloader.start_process_func()
    
    # output_directory 가져오기
    output_directory = sound_downloader.output_directory

    # output_directory의 파일들 불러오기
    loaded_files = soundController.load_files_from_directory(output_directory)

    if loaded_files:
        print("Files in the directory:")
        for file_path in loaded_files:
            print(file_path)
    else:
        print("No files loaded.")