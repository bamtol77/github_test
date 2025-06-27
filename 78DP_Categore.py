import pandas as pd
import os

# 엑셀 파일 경로 설정
file_path = 'data.xlsx'
file_root, file_ext = os.path.splitext(file_path)
# 처리된 파일을 저장할 새 경로: 원본 파일명_1.확장자
new_file_path = f"{file_root}_1{file_ext}"

# --- 원본 시트 불러오기 및 초기 데이터 처리 ---
try:
    df = pd.read_excel(file_path, sheet_name='rg-grid')
except FileNotFoundError:
    print(f"오류: '{file_path}' 파일을 찾을 수 없습니다. 파일 경로를 확인해주세요.")
    exit()
except ValueError:
    print(f"오류: '{file_path}' 파일에 'rg-grid' 시트가 없습니다. 시트 이름을 확인해주세요.")
    exit()

# 필요한 컬럼들을 문자형식으로 변환 (NaN 안전하게 처리)
# .astype(str) 전에 .fillna('')를 사용하여 NaN을 빈 문자열로 대체하면 더 안전합니다.
df['규약번호'] = df['규약번호'].fillna('').astype(str)
df['RCS'] = df['RCS'].fillna('').astype(str)
df['RA하한'] = df['RA하한'].fillna('').astype(str)
df['출강목표'] = df['출강목표'].fillna('').astype(str)
df['사내보증번호'] = df['사내보증번호'].fillna('').astype(str)
df['SBT등급'] = df['SBT등급'].fillna('').astype(str)


# '소재구분' 시트 생성을 위한 빈 DataFrame 초기화
required_columns = [
    '순번', '코일번호', '두께', '폭', '길이', 'SS온도', 'RCS온도', 'RA하한',
    '내부산화', 'SBT등급', 'SBT시험', '잔공정', '출강목표', '사내보증번호',
    '도금량', '규약번호', '고객사명'
]
df_material_category = pd.DataFrame(columns=['구분'] + required_columns)

# --- 각 조건에 따른 데이터 필터링 및 '소재구분' 데이터프레임에 추가 ---

# '규약번호'가 '780DH'를 포함하는 경우
condition_780dh = df['규약번호'].str.contains('780DH', na=False)
if not df[condition_780dh].empty:
    temp_df = df[condition_780dh][required_columns].copy()
    temp_df.insert(0, '구분', '780DH')
    df_material_category = pd.concat([df_material_category, temp_df], ignore_index=True)

# '규약번호'가 '780CP'를 포함하는 경우
condition_780cp = df['규약번호'].str.contains('780CP', na=False)
if not df[condition_780cp].empty:
    temp_df = df[condition_780cp][required_columns].copy()
    temp_df.insert(0, '구분', '780CP')
    df_material_category = pd.concat([df_material_category, temp_df], ignore_index=True)

# '규약번호'가 '780DP'를 포함하고, 출강목표가 'C100'으로 시작하고, 'RCS'=='450'인 경우
condition_low_si_780dp = (
    df['규약번호'].str.contains('780DP', na=False) &
    df['출강목표'].str.startswith('C100', na=False) &
    (df['RCS'] == '450')
)
if not df[condition_low_si_780dp].empty:
    temp_df = df[condition_low_si_780dp][required_columns].copy()
    temp_df.insert(0, '구분', '저 Si_780DP')
    df_material_category = pd.concat([df_material_category, temp_df], ignore_index=True)

# '규약번호'가 '780DP'를 포함하고, 출강목표가 'C070'으로 시작하고, 'RCS'=='500'인 경우
condition_high_si_780dp = (
    df['규약번호'].str.contains('780DP', na=False) &
    df['출강목표'].str.startswith('C070', na=False) &
    (df['RCS'] == '500')
)
if not df[condition_high_si_780dp].empty:
    temp_df = df[condition_high_si_780dp][required_columns].copy()
    temp_df.insert(0, '구분', '고 Si_780DP')
    df_material_category = pd.concat([df_material_category, temp_df], ignore_index=True)

# '규약번호'가 '980CP'를 포함하는 경우
condition_980cp = df['규약번호'].str.contains('980CP', na=False)
if not df[condition_980cp].empty:
    temp_df = df[condition_980cp][required_columns].copy()
    temp_df.insert(0, '구분', '980CP_W반곡 발생재')
    df_material_category = pd.concat([df_material_category, temp_df], ignore_index=True)

# '규약번호'가 '980XF'를 포함하고, '사내보증번호'가 '980X'를 포함하는 경우
condition_980xf = (
    df['규약번호'].str.contains('980XF', na=False) &
    df['사내보증번호'].str.contains('980X', na=False)
)
if not df[condition_980xf].empty:
    temp_df = df[condition_980xf][required_columns].copy()
    temp_df.insert(0, '구분', '980XF')
    df_material_category = pd.concat([df_material_category, temp_df], ignore_index=True)

# '규약번호'가 '980DP'를 포함하고, '출강목표' 5~7 번째 자리에 '250'으로 표시된 경우
condition_high_ys_980dp = (
    df['규약번호'].str.contains('980DP', na=False) &
    df['출강목표'].str[4:7].str.contains('250', na=False)
)
if not df[condition_high_ys_980dp].empty:
    temp_df = df[condition_high_ys_980dp][required_columns].copy()
    temp_df.insert(0, '구분', '고YS_980DP')
    df_material_category = pd.concat([df_material_category, temp_df], ignore_index=True)

# '규약번호'가 '980DP'를 포함하고, '출강목표' 5~7 번째 자리에 '230'으로 표시된 경우
condition_low_ceq_980dp = (
    df['규약번호'].str.contains('980DP', na=False) &
    df['출강목표'].str[4:7].str.contains('230', na=False)
)
if not df[condition_low_ceq_980dp].empty:
    temp_df = df[condition_low_ceq_980dp][required_columns].copy()
    temp_df.insert(0, '구분', '저 CEQ_980DP')
    df_material_category = pd.concat([df_material_category, temp_df], ignore_index=True)

# '규약번호'가 '980DP'를 포함하고, 'RA하한'에 숫자나 문자가 있고, '출강목표' 5~7 번째 자리에 '230'으로 표시된 경우
condition_no_work_low_ceq_980dp = (
    df['규약번호'].str.contains('980DP', na=False) &
    (df['RA하한'].str.len() > 0) & # RA하한이 비어있지 않은 경우
    df['출강목표'].str[4:7].str.contains('230', na=False)
)
if not df[condition_no_work_low_ceq_980dp].empty:
    temp_df = df[condition_no_work_low_ceq_980dp][required_columns].copy()
    temp_df.insert(0, '구분', '작업불가_저 CEQ_980DP')
    df_material_category = pd.concat([df_material_category, temp_df], ignore_index=True)

# --- 'A코팅 작업여부' 항목 생성 및 조건 적용 ---
df_material_category['A코팅 작업여부'] = '' # 컬럼 생성 및 초기화

# '구분'이 '저 Si_780DP'이고, 'SBT등급'이 '5'이거나 비어있는 경우
condition_a_coating_optional = (
    (df_material_category['구분'] == '저 Si_780DP') &
    ((df_material_category['SBT등급'] == '5') | (df_material_category['SBT등급'] == ''))
)
df_material_category.loc[condition_a_coating_optional, 'A코팅 작업여부'] = 'A코팅_무시가능'

# '구분'이 '저 Si_780DP'이고, 'SBT등급'이 '1~4'인 경우
condition_a_coating_required_low_si = (
    (df_material_category['구분'] == '저 Si_780DP') &
    (df_material_category['SBT등급'].isin(['1', '2', '3', '4']))
)
df_material_category.loc[condition_a_coating_required_low_si, 'A코팅 작업여부'] = 'A코팅_필수'

# '구분'이 '고 Si_780DP'인 경우
condition_a_coating_required_high_si = (df_material_category['구분'] == '고 Si_780DP')
df_material_category.loc[condition_a_coating_required_high_si, 'A코팅 작업여부'] = 'A코팅_필수'

# --- 엑셀 파일 저장 및 조건부 서식 적용 ---

# '구분' 색상 매핑 딕셔너리 정의
color_map = {
    '780DH': '#ABEBC6',
    '780CP': '#82E0AA',
    '저 Si_780DP': '#45B39D',
    '고 Si_780DP': '#138D75',
    '980CP_W반곡 발생재': '#F8C471',
    '980XF': '#F5B041',
    '980DP': '#D68910',
    '고YS_980DP': '#EDBB99',
    '저 CEQ_980DP': '#E59866',
    '작업불가_저 CEQ_980DP': '#CA6F1E'
}

    # '소재구분' 시트에 조건부 서식 적용
    workbook = writer.book
    worksheet = writer.sheets['소재구분']

    # '구분' 열의 인덱스 (엑셀 A열)
    category_col_idx = 0

    # 'A코팅 작업여부' 열의 인덱스 찾기
    try:
        a_coating_col_idx = df_material_category.columns.get_loc('A코팅 작업여부')
    except KeyError:
        print("경고: 'A코팅 작업여부' 컬럼을 찾을 수 없습니다. 관련 서식을 적용할 수 없습니다.")
        a_coating_col_idx = -1 # 찾지 못했을 경우 유효하지 않은 인덱스 할당

    # 조건부 서식 적용 범위 설정 (헤더 제외)
    start_row = 1 # 엑셀에서 데이터는 1행(인덱스 0)부터 시작, 헤더(0행) 다음이므로 1
    end_row = len(df_material_category) # 데이터의 마지막 행 (데이터프레임 길이 = 마지막 인덱스 + 1)

    # '구분' 열에 대한 배경색 조건부 서식 적용
    for category_name, hex_color in color_map.items():
        cell_format = workbook.add_format({'bg_color': hex_color})
        worksheet.conditional_format(
            start_row, category_col_idx, end_row, category_col_idx,
            {
                'type': 'text_string', # 정확한 문자열 일치를 위해 'text_string' 사용
                'criteria': '==',
                'value': category_name,
                'format': cell_format
            }
        )

    # 'A코팅 작업여부' 열에 대한 글씨색 조건부 서식 적용
    if a_coating_col_idx != -1: # 'A코팅 작업여부' 컬럼이 존재할 경우에만 서식 적용
        red_font_format = workbook.add_format({'font_color': '#ff0000'})
        worksheet.conditional_format(
            start_row, a_coating_col_idx, end_row, a_coating_col_idx,
            {
                'type': 'text_string',
                'criteria': '==',
                'value': 'A코팅_필수',
                'format': red_font_format
            }
        )

print(f"데이터 처리 및 엑셀 서식 변경이 완료되었습니다. 결과는 '{new_file_path}' 파일에 저장되었습니다.")

# 결과를 새로운 엑셀 파일로 저장
with pd.ExcelWriter(new_file_path, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='rg-grid', index=False) # 원본 시트 유지
    df_material_category.to_excel(writer, sheet_name='소재구분', index=False) # 새로운 '소재구분' 시트 저장
