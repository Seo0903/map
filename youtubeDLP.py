import os
import pandas as pd
import yt_dlp
from pydub import AudioSegment
import subprocess as sp

class SoundDownloader():
    def __init__(self):
        self.efn = None
        self.output_directory = None
        self.sud = None
        self.range1 = None
        self.range2 = None

    def process_values(self, ui_values):
        self.efn = ui_values["efn"]
        self.output_directory = ui_values["output_directory"]
        self.sud = ui_values["sud"]
        self.range1 = ui_values["range1"]
        self.range2 = ui_values["range2"]

    def start_process_func(self):
        if self.output_directory is None:
            print("No output directory selected. Please select an output directory.\n")
            return
        current_directory = os.path.dirname(__file__)
        exfile = pd.read_excel(self.efn, sheet_name='Sheet1')
        if (self.range2 == 'M' or self.range2 == 'm'):
            colnum = exfile['링크'].count()
        else:
            colnum = int(self.range2) - 1
        print("목록 개수 : " + str(colnum))
        for i in range(int(self.range1) - 1, int(colnum)):
            if os.path.exists(os.path.join(self.output_directory, str(i + 1) + '.' + exfile.at[i, '작품명'] + '.ogg')):
                os.remove(os.path.join(self.output_directory, str(i + 1) + '.' + exfile.at[i, '작품명'] + '.ogg'))
            # 유튜브 전용 인스턴스 생성
            ydl_opts = {
                'format': 'bestaudio/best',
                'extractaudio': True,
                'audioformat': 'aac',
                'outtmpl': str(i + 1) + "." + exfile.at[i, '작품명'] + '.aac',
            }
            chk = 0
            while chk != 1:
                try:
                    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                        print(ydl.download(exfile.at[i, '링크']))
                except (Exception):
                    print('sibar')
                    break
                else:
                    # 특정 영상 다운로드
                    chk = 1
                    print("다운로드 완료 ({}/{})".format(i + 1, colnum))

            sound = AudioSegment.from_file(current_directory + "/" + str(i + 1) + "." + exfile.at[i, '작품명'] + ".aac")
            StrtMin = exfile.at[i, '시작분']
            StrtSec = exfile.at[i, '시작초']
            EndMin = exfile.at[i, '끝분']
            EndSec = exfile.at[i, '끝초']
            StrtTime = StrtMin * 60 * 1000 + StrtSec * 1000
            EndTime = EndMin * 60 * 1000 + EndSec * 1000
            extract = sound[StrtTime:EndTime]
            extract.export(current_directory + "/" + str(i + 1) + "." + exfile.at[i, '작품명'] + ".mp3", format="mp3")
            os.remove(current_directory + "/" + str(i + 1) + "." + exfile.at[i, '작품명'] + ".aac")
            print("자르기 완료 ({}/{})".format(i + 1, colnum))

            # 저장될 파일 경로를 지정
            output_file_path = os.path.join(self.output_directory, f"{i + 1}.{exfile.at[i, '작품명']}.ogg")

            # 음질을 지정한 값으로 조절
            result = sp.Popen(['ffmpeg', '-i',
                               os.path.join(current_directory, str(i + 1) + '.' + exfile.at[i, '작품명'] + '.mp3'),
                               '-af',
                               'loudnorm=I=-18:LRA=11:TP=-2:measured_I=-20.26:measured_LRA=14.1:measured_TP=-9.22:measured_thresh=-31.20:offset=1.28:linear=true, volume='
                               + str(self.sud), output_file_path], stdout=sp.PIPE, stderr=sp.PIPE)
            out, err = result.communicate()
            exitcode = result.returncode
            if exitcode != 0:
                print(exitcode, out.decode('utf8'), err.decode('utf8'))
            else:
                print("음질 작업 완료 ({}/{})".format(i + 1, colnum))
                os.remove(
                    os.path.join(current_directory, str(i + 1) + '.' + exfile.at[i, '작품명'] + '.mp3'))
