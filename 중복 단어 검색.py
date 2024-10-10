import pandas as pd
import collections
import os
import re

# 엑셀 파일을 CSV로 변환하는 함수
def excel_to_csv(excel_file_path, output_dir):
    # 엑셀 파일의 모든 시트 읽기
    xls = pd.ExcelFile(excel_file_path)
    
    # 각 시트를 반복하여 CSV로 저장
    csv_file_paths = []  # 생성된 CSV 파일 경로를 저장할 리스트
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        csv_file_path = os.path.join(output_dir, f'{sheet_name}.csv')  # 시트 이름으로 CSV 파일 경로 생성
        df.to_csv(csv_file_path, index=False, encoding="utf-8")
        print(f'{csv_file_path}로 변환되었습니다.')
        csv_file_paths.append(csv_file_path)  # 생성된 CSV 파일 경로 추가

    return csv_file_paths  # 생성된 CSV 파일 경로 리스트 반환

# 중복 문자 및 숫자를 찾고 결과를 리스트로 반환하는 함수
def find_duplicate_min_five_chars(csv_file_path):
    # 중복된 5글자 이상의 단어 및 숫자를 저장할 카운터
    counter = collections.Counter()
    
    # CSV 파일에서 데이터 읽기
    with open(csv_file_path, "r", encoding="utf-8") as file:
        for line in file:  # 파일을 한 줄씩 읽기
            # 5글자 이상 길이의 단어 및 숫자 찾기 (정규 표현식 사용)
            words = re.findall(r'\b\w{5,}\b', line)  # 5글자 이상 단어 찾기
            counter.update(words)  # 카운터에 단어 추가

    # 중복된 5글자 이상 길이의 결과를 리스트에 저장
    duplicates = []
    for char, count in counter.items():
        if count > 1:  # 중복된 항목만 확인
            duplicates.append((char, count))  # (문자/숫자, 빈도수) 형태로 저장
            print(f'"{char}"이(가) {count}번 중복되었습니다.')
    
    return duplicates

# 엑셀 파일 경로와 변환될 CSV 파일이 저장될 디렉토리
excel_file_path = "C:/Users/namgeol/Desktop/Shop.xlsx"
output_dir = "C:/Users/namgeol/Desktop"  # CSV 파일이 저장될 디렉토리

# 엑셀 파일을 CSV 파일로 변환하고 생성된 CSV 파일 경로 리스트 받기
csv_file_paths = excel_to_csv(excel_file_path, output_dir)

# 중복 5글자 이상 결과를 저장할 리스트
all_duplicates = []

# 변환된 CSV 파일에서 중복 5글자 이상 찾기
for csv_file_path in csv_file_paths:  # 생성된 CSV 파일 경로 리스트 반복
    duplicates = find_duplicate_min_five_chars(csv_file_path)
    all_duplicates.extend(duplicates)  # 모든 중복 5글자 이상 결과를 리스트에 추가

# 모든 중복 5글자 이상 결과를 하나의 Excel 파일의 시트로 저장
results_df = pd.DataFrame(all_duplicates, columns=["Character", "Count"])
results_excel_path = os.path.join(output_dir, "results.xlsx")

# ExcelWriter를 사용하여 결과를 단일 시트에 저장
with pd.ExcelWriter(results_excel_path, engine='openpyxl') as writer:
    results_df.to_excel(writer, index=False, sheet_name='Duplicates')

print(f'중복 5글자 이상 결과가 {results_excel_path}에 저장되었습니다.')

# 생성된 CSV 파일 삭제
for csv_file_path in csv_file_paths:
    os.remove(csv_file_path)  # CSV 파일 삭제
    print(f'{csv_file_path}가 삭제되었습니다.')
