import pandas as pd
import os
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook
import re

# 사용자 바탕화면 경로 자동 탐지
desktop_dir = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')

# 엑셀 파일 경로 (사용자 바탕화면에 Shop.xlsx 파일이 있다고 가정)
excel_file_path = os.path.join(desktop_dir, "Shop.xlsx")

# CSV 파일 및 결과 파일이 저장될 디렉토리 (사용자 바탕화면)
output_dir = desktop_dir

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


# EndDate로 시작하는 열에서 날짜와 숫자만 남기는 함수
def find_enddate_info(csv_file_path):
    # 정보를 저장할 리스트
    enddate_info = []

    # CSV 파일에서 데이터 읽기
    df = pd.read_csv(csv_file_path, encoding="utf-8")
    
    # 데이터 확인 (디버깅용)
    print("CSV 데이터 샘플:")
    print(df.head())  # 데이터의 첫 5줄을 출력
    
    # EndDate로 시작하는 열 찾기
    enddate_columns = [col for col in df.columns if col.startswith('EndDate')]
    
    # 찾은 EndDate 열을 출력 (디버깅용)
    print("찾은 EndDate 열들:", enddate_columns)
    
    if not enddate_columns:
        print("EndDate로 시작하는 열이 없습니다.")
        return enddate_info
    
    # 날짜 또는 숫자 데이터가 포함된 열과 앞의 열, 첫 번째와 두 번째 열 정보를 리스트화
    for col in enddate_columns:
        for idx, date_value in df[col].items():
            # 날짜 데이터 필터링
            try:
                # 날짜 형식 확인 및 변환
                date_value = pd.to_datetime(date_value, errors='raise')  # 오류 발생 시 예외 처리
            except:
                # 숫자 데이터만 필터링
                if isinstance(date_value, str) and not re.match(r'^\d+$', date_value):
                    continue  # 숫자가 아닌 경우 건너뜀
            
            if pd.notnull(date_value):  # 유효한 날짜나 숫자 값이 존재하면
                # 앞 열의 정보, 첫 번째, 두 번째 열의 정보 추가
                previous_col_value = df.iloc[idx, df.columns.get_loc(col) - 1]  # 날짜 앞 열의 값
                first_col_value = df.iloc[idx, 0]  # 첫 번째 열의 값
                second_col_value = df.iloc[idx, 1]  # 두 번째 열의 값
                
                # 디버깅을 위한 출력
                print(f'Index: {idx}, 날짜/숫자 정보: {date_value}, 앞 열의 정보: {previous_col_value}, 첫 번째 열: {first_col_value}, 두 번째 열: {second_col_value}')
                
                enddate_info.append([first_col_value, second_col_value, previous_col_value, date_value])

    return enddate_info

# 엑셀 파일을 CSV 파일로 변환하고 생성된 CSV 파일 경로 리스트 받기
csv_file_paths = excel_to_csv(excel_file_path, output_dir)

all_enddate_info = []

# 변환된 CSV 파일에서 EndDate 열 관련 정보 찾기
for csv_file_path in csv_file_paths:  # 생성된 CSV 파일 경로 리스트 반복
    enddate_info = find_enddate_info(csv_file_path)
    all_enddate_info.extend(enddate_info)  # 모든 EndDate 관련 정보를 리스트에 추가

# EndDate 관련 정보를 엑셀 파일의 시트로 저장하고 열 너비와 행 높이를 자동으로 설정하는 함수
def save_to_excel_with_auto_width_and_height(df, output_path, sheet_name='Sheet1'):
    # DataFrame을 엑셀 파일로 저장
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        
        # 저장된 엑셀 파일을 불러오기
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        # 열 너비를 자동으로 설정
        for col in worksheet.columns:
            max_length = 0
            column = col[0].column_letter  # 열의 알파벳 (A, B, C 등)
            for cell in col:
                try:  # 셀의 값이 None인 경우를 방지
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            adjusted_width = (max_length + 2)  # 여유를 두고 열 너비 설정
            worksheet.column_dimensions[column].width = adjusted_width
        
        # 행 높이를 자동으로 설정
        for row in worksheet.iter_rows():
            max_height = 0
            for cell in row:
                try:
                    if cell.value:
                        max_height = max(max_height, len(str(cell.value)) // 10 + 15)  # 높이 설정 기준 (여유를 두고 15 추가)
                except:
                    pass
            worksheet.row_dimensions[row[0].row].height = max_height

    print(f'{output_path}에 저장 완료. 열 너비 및 행 높이 자동 설정됨.')

# EndDate 관련 정보를 내림차순으로 정렬하고 엑셀 파일로 저장
results_df = pd.DataFrame(all_enddate_info, columns=["상품 tid", "상품명", "StartDate", "EndDate"])

# EndDate 열을 기준으로 내림차순 정렬
results_df = results_df.sort_values(by="EndDate", ascending=False)

# 정렬된 데이터를 저장할 엑셀 파일 경로
results_excel_path = os.path.join(output_dir, "shop결과.xlsx")

# EndDate 관련 정보를 엑셀 파일로 저장하며 열 너비와 행 높이 자동 조정
save_to_excel_with_auto_width_and_height(results_df, results_excel_path, sheet_name='EndDate_Info')
print(f'EndDate 관련 정보가 {results_excel_path}에 저장되었습니다.')

# 생성된 CSV 파일 삭제
for csv_file_path in csv_file_paths:
    os.remove(csv_file_path)  # CSV 파일 삭제
    print(f'{csv_file_path}가 삭제되었습니다.')
