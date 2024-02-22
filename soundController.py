import os
from pydub import AudioSegment

def load_files_from_directory(output_directory):
    if not os.path.exists(output_directory):
        print(f"The specified directory '{output_directory}' does not exist.")
        return None

    # 파일 리스트 불러오기
    all_files = os.listdir(output_directory)

    # ogg 확장자를 가진 파일만 선택
    ogg_files = [file for file in all_files if file.lower().endswith('.ogg')]

    return ogg_files

def combine_files(file_list, output_directory, output_file_name):
    if not file_list:
        print("No files to combine.")
        return

    combined = AudioSegment.from_file(os.path.join(output_directory, file_list[0]))

    for file_name in file_list[1:]:
        sound = AudioSegment.from_file(os.path.join(output_directory, file_name))
        combined += sound

    # 결과 파일 저장
    combined.export(os.path.join(output_directory, output_file_name), format="ogg")

    print(f"Combined files saved as {output_file_name}")

# 사용 예시
output_directory = "/path/to/your/directory"
ogg_files = load_files_from_directory(output_directory)

if ogg_files:
    print("OGG files in the directory:")
    for file_name in ogg_files:
        print(file_name)

    # 파일 합치기
    combine_files(ogg_files, output_directory, "combined_result.ogg")
else:
    print("No OGG files found.")
