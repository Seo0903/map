import os
from pydub import AudioSegment

save_directory = None

def load_files_from_directory(output_directory):
    save_directory = output_directory
    if not os.path.exists(output_directory):
        print(f"'{output_directory}' 디렉토리가 존재하지 않습니다.")
        return None

    # 파일 목록 불러오기
    all_files = os.listdir(output_directory)

    # .ogg 확장자를 가진 파일만 선택
    ogg_files = [file for file in all_files if file.lower().endswith('.ogg')]

    return ogg_files

def combine_files(file_list, output_directory, output_file_name):
    if not file_list:
        print("합칠 파일이 없습니다.")
        return

    combined = AudioSegment.from_file(os.path.join(output_directory, file_list[0]))

    for file_name in file_list[1:]:
        sound = AudioSegment.from_file(os.path.join(output_directory, file_name))
        combined += sound

    # 결과 파일 저장
    combined.export(os.path.join(output_directory, output_file_name), format="ogg")

    print(f"결합된 파일을 '{output_file_name}'으로 저장하였습니다.")

# 사용 예시
output_directory = "{save_directory}/combine"

# 디렉토리가 존재하지 않으면 생성
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

ogg_files = load_files_from_directory(output_directory)

if ogg_files:
    print("디렉토리에 있는 OGG 파일 목록:")
    for i, file_name in enumerate(ogg_files):
        print(file_name)
        # 파일명에 번호 추가
        new_file_name = f"{i+1}_{file_name}"
        ogg_files[i] = new_file_name

    # 파일 합치기
    combine_files(ogg_files, output_directory, "combined_result.ogg")
else:
    print("OGG 파일을 찾을 수 없습니다.")
