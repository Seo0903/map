import openpyxl

# ㄱ ~ ㅣ 자음 모음 리스트
# ㄱ ~ ㅣ 자음 모음 리스트
chosungs = ["ㄱ","ㄲ","ㄴ","ㄷ","ㄸ","ㄹ","ㅁ","ㅂ","ㅃ","ㅅ","ㅆ","ㅇ","ㅈ","ㅉ","ㅊ","ㅋ","ㅌ","ㅍ","ㅎ"]
jungsungs = ["ㅏ","ㅐ","ㅑ","ㅒ","ㅓ","ㅔ","ㅕ","ㅖ","ㅗ","ㅘ","ㅙ","ㅚ","ㅛ","ㅜ","ㅝ","ㅞ","ㅟ","ㅠ","ㅡ","ㅢ","ㅣ"]
jongsungs = ["","ㄱ","ㄲ","ㄳ","ㄴ","ㄵ","ㄶ","ㄷ","ㄹ","ㄺ","ㄻ","ㄼ","ㄽ","ㄾ","ㄿ","ㅀ","ㅁ","ㅂ","ㅄ","ㅅ","ㅆ","ㅇ","ㅈ","ㅊ","ㅋ","ㅌ","ㅍ","ㅎ"]

# 한글 음절 자음 모음 분리 함수
def separate_kor_char(text):
    result = []
    for char in text:
        if ord('가') <= ord(char) <= ord('힣'):
            char_code = ord(char) - ord('가')
            jong = char_code % 28
            jung = ((char_code - jong) // 28) % 21
            cho = (((char_code - jong) // 28) - jung) // 21
            result.append(chosungs[cho])
            result.append(jungsungs[jung])
            if jong > 0:
                result.append(jongsungs[jong])
        else:
            result.append(char)
    return result
# 1. 액셀 파일 열기
workbook = openpyxl.load_workbook('E:/mapList/5. 한소절/LiteExslCutter/한소절.xlsx')
worksheet = workbook['Sheet1']
ws = workbook.active

# 데이터를 저장할 이중 리스트 생성
data = []
han = []


# G열의 2번째 행부터 마지막 행까지 데이터를 읽어와 이중 리스트에 저장
for row in worksheet.iter_rows(min_row=2, min_col=7, values_only=True):
    # None 값을 필터링하여 데이터 저장
    filtered_row = [value for value in row if value is not None]
    if filtered_row:
        split_row = []
        split_han = []  # 추가된 부분
        for value in filtered_row:
            if isinstance(value, str):
                split_value = value.split("/")
                if len(split_value) >= 2:
                    split_row.append(split_value[0])
                    split_han.extend(split_value[1:])  # 수정된 부분
                else:
                    split_row.append(value)
            else:
                split_row.append(value)
        data.append(split_row)
        han.append(split_han)  # 추가된 부분

max_row = ws.max_row
for row in range(2, max_row+1):
    if ws.cell(row=row, column=7).value is None:
        max_row = row - 1
        break

# 데이터가 있는 행만큼 반복하면서 view 리스트에 추가
view = []
for row in range(2, max_row+1):
    row_data = []
    for col in range(7, 27):
        if ws.cell(row=row, column=col).value is None:
            row_data.append(-2)
        else:
            row_data.append(-1)
    view.append(row_data + [99])


# chk 리스트 생성 및 중복없이 data 리스트에 있는 모든 값을 추가
chk = []
for row in data:
    for value in row:
        if value not in chk:
            chk.append(value)

chk_dict = {item: i+1 for i, item in enumerate(chk)}
ind = [[chk_dict[item] for item in row] + [-1] for row in data]
cti = list(set(chk_dict[item] for row in data for item in row))

chk_dict2 = dict(zip(chk, cti))

with open('text.txt', 'w') as f:
    for key, value in chk_dict2.items():
        f.write(f"{key}: {value+1}\n")
with open("text.txt", "r") as file:
    text_list = [line.strip().replace("\n", "") for line in file.readlines()]
    
my_dict = {}
for text in text_list:
    # val: key 구조라 가정
    key = text.split(":")[-1]
    key = key.strip()
    val = ":".join(*text.split(":")[:-1])
    val = val.strip()
    if my_dict.get(key) == None:
        my_dict[key] = []
    my_dict[key].append(val)


underB = []
for i in range(len(data)):
    row = []
    for j in range(len(data[i])):
        cell = ""
        for k in range(len(data[i][j])):
            if data[i][j][k] == " ":
                cell += "  "
            else:
                cell += "_ "
        row.append(cell.strip())
    underB.append(row)

replace_dict = {
    "ㅗ": ["오", "우"],
    "ㅛ": ["오", "우"],
    "ㅜ": ["우"],
    "ㅠ": ["우"],
    "ㅑ": ["아"],
    "ㅏ": ["아"],
    "ㅘ": ["아"],
    "ㅖ": ["에", "이"],
    "ㅔ": ["에", "이"],
    "ㅣ": ["이"],
}

text_list = []

from copy import deepcopy
# key는 idx, val은 문자
for key, val_list in deepcopy(my_dict).items():
    for val in val_list:
        if "-" not in val:
            text_list.append(f"{val}: {key}\n")
            continue
        index_dict = {}
        for i, v in enumerate(val):
            if v == "-":
                index_dict[i] = replace_dict(separate_kor_char(val[i-1]))
        temp_list = [""]
        for i, v in enumerate(val):
            if i in index_dict:
                temp_list = [t+r for t in temp_list for r in replace_dict[index_dict[i]]]
            else:
                temp_list = [t+v for t in temp_list]
        for val2 in temp_list:
            text_list.append(f"{val2}: {key}\n")
    
seen_values = set()
result2 = []
for item in text_list:
    left_value = item.split(":")[0]
    if left_value not in seen_values:
        result2.append(item)
        seen_values.add(left_value)

text_list = result2

seen_values = set()
result3 = []
for item in text_list:
    left_value = item.split(":")[0]
    if left_value not in seen_values:
        result3.append(item)
        seen_values.add(left_value)

text_list = result3

for i in range(len(data)):
    for j in range(len(data[i])):
        data[i][j] = 'Db("' + str(data[i][j]) + '")'
    for k in range(len(han[i])):
        han[i][k] = 'Db("' + str(han[i][k]) + '")'
    for m in range(len(underB[i])):
        underB[i][m] = 'Db("' + str(underB[i][m]) + '")'

with open('output.txt', 'w') as f:
    f.write('const dataA = \n[0,')
    for row in data:
        f.write('[' + ', '.join(row) + '],\n')
    f.write('];\n')
    f.write('\n')
    f.write('const han = \n[0,')
    for row in han:
        f.write('[' + ', '.join(row) + '],\n')
    f.write('];\n')
    f.write('\n')
    f.write('const underB = \n[0,')
    for row in underB:
        f.write('[' + ', '.join(row) + '],\n')
    f.write('];\n')
    f.write('\n')
    f.write('const ind = \n[0,')   
    for row in ind:
        f.write(str(row) + ',\n')
    f.write('];\n')
    f.write('\n')
    f.write('const view = \n[0,')
    for row in view:
        f.write(str(row) + ',\n')
    f.write('];\n')
    f.write('\n')
    for row in text_list:
        f.write(str(row) + '\n')
    f.write('\n')
    f.write('\n')