
import pandas as pd
import os

def process_and_categorize_data(file_path):
    """
    엑셀 파일을 로드하고 'ag-grid' 시트의 데이터를 읽어온 후,
    조건에 따라 데이터를 필터링하여 5개의 새로운 시트에 저장합니다.
    '이송재정리' 시트는 여러 조건을 만족하는 데이터를 옆으로 이어서 저장합니다.

    Args:
        file_path (str): 처리할 엑셀 파일의 경로.
    """
    try:
        # 1. 엑셀 파일 로딩 및 'ag-grid' 시트 선택
        # Loading the Excel file and selecting the 'ag-grid' sheet
        df = pd.read_excel(file_path, sheet_name='ag-grid')
        print(f"'{file_path}' 파일에서 'ag-grid' 시트를 성공적으로 로드했습니다.")
        print(f"원본 데이터 행 수: {len(df)}")
        print("---")

        # 공통으로 추출할 컬럼 리스트
        # List of common columns to extract
        target_columns = [
            '순번', '코일번호', '두께', '폭', '길이', '중량',
            '사내보증번호(구)', '(현)저장위치', '후처리', '고객사명', '고객사'
        ]

        # 저장할 파일명 생성 (원본 파일명에 _1 붙이기)
        # Creating the filename for saving (appending _1 to the original filename)
        directory, original_filename = os.path.split(file_path)
        name, ext = os.path.splitext(original_filename)
        new_filename = f"{name}_1{ext}"
        save_path = os.path.join(directory, new_filename)

        # ExcelWriter 객체를 사용하여 여러 시트에 데이터 쓰기
        # Using ExcelWriter to write data to multiple sheets
        with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:

            ## 1. '3동간_이적' 시트 처리
            # Processing '3동간_이적' sheet
            print("1. '3동간_이적' 시트 데이터 필터링 중...")
            condition_3donggan = (
                df['(현)저장위치'].isin(['KA1', 'KA2', 'KA3']) &
                df['코일번호'].astype(str).str.startswith('CR', na=False)
            )
            df_3donggan = df[condition_3donggan][target_columns]
            df_3donggan.to_excel(writer, sheet_name='3동간_이적', index=False)
            print(f" - '{len(df_3donggan)}' 행 저장 완료.")

            ## 2. '통과공정이상재' 시트 처리
            # Processing '통과공정이상재' sheet
            print("2. '통과공정이상재' 시트 데이터 필터링 중...")
            # 'CAL,CGL가능동과공정' 컬럼이 문자열이 아닐 경우를 대비해 .astype(str) 적용
            condition_process_issue = ~(
                df['CAL,CGL가능동과공정'].astype(str).str.contains('7', na=False)
            )
            df_process_issue = df[condition_process_issue][target_columns]
            df_process_issue.to_excel(writer, sheet_name='통과공정이상재', index=False)
            print(f" - '{len(df_process_issue)}' 행 저장 완료.")

            ## 3. '후처리이상재' 시트 처리
            # Processing '후처리이상재' sheet
            print("3. '후처리이상재' 시트 데이터 필터링 중...")
            exclude_post_processing = ['XX', 'LP', 'NC']
            # '후처리' 컬럼이 문자열이 아닐 경우를 대비해 .astype(str) 적용
            condition_post_process_issue = ~(
                df['후처리'].astype(str).isin(exclude_post_processing)
            )
            df_post_process_issue = df[condition_post_process_issue][target_columns]
            df_post_process_issue.to_excel(writer, sheet_name='후처리이상재', index=False)
            print(f" - '{len(df_post_process_issue)}' 행 저장 완료.")

            ## 4. '자가재' 시트 처리
            # Processing '자가재' sheet
            print("4. '자가재' 시트 데이터 필터링 중...")
            # '고객사' 컬럼이 문자열이 아닐 경우를 대비해 .astype(str) 적용
            condition_own_material = (
                df['고객사'].astype(str) == 'KKKK'
            )
            df_own_material = df[condition_own_material][target_columns]
            df_own_material.to_excel(writer, sheet_name='자가재', index=False)
            print(f" - '{len(df_own_material)}' 행 저장 완료.")

            ## 5. '이송재정리' 시트 처리 (복합 조건 및 옆으로 이어붙이기)
            # Processing '이송재정리' sheet (complex conditions and appending horizontally)
            print("5. '이송재정리' 시트 데이터 필터링 및 병합 중...")

            # 기본 조건 정의
            base_condition = (
                (df['Status'].astype(str) == 'N') &
                (df['(현)진도'].astype(str) == 'H')
            )

            # 각 조건별 데이터프레임 생성
            df_cp = df[base_condition & df['코일번호'].astype(str).str.startswith('CP', na=False)][target_columns]
            df_cq = df[base_condition & df['코일번호'].astype(str).str.startswith('CQ', na=False)][target_columns]
            df_cr = df[base_condition & df['코일번호'].astype(str).str.startswith('CR', na=False)][target_columns]
            df_cs = df[base_condition & df['코일번호'].astype(str).str.startswith('CS', na=False)][target_columns]

            # 컬럼 이름 충돌 방지를 위해 각 데이터프레임의 컬럼에 접두사 추가
            # Add prefixes to column names to avoid collision when concatenating horizontally
            df_cp = df_cp.add_suffix('_CP')
            df_cq = df_cq.add_suffix('_CQ')
            df_cr = df_cr.add_suffix('_CR')
            df_cs = df_cs.add_suffix('_CS')

            # 데이터를 옆으로 이어 붙이기 (outer join 사용하여 모든 행 포함)
            # Concatenate dataframes horizontally using outer join to include all rows
            # '순번' 컬럼을 기준으로 합치면 안되므로, 인덱스를 재설정하거나 join하지 않고 단순 concat
            # Here, we'll write them sequentially using startcol parameter
            
            # 각 데이터프레임을 별도의 시작 열에 씁니다.
            # Write each dataframe to a separate starting column.
            # 이송재정리 시트는 ExcelWriter의 startrow, startcol 파라미터를 활용하여
            # 원하는 위치에 데이터를 쓸 수 있습니다.
            # 첫 번째 데이터프레임 (CP)
            if not df_cp.empty:
                df_cp.to_excel(writer, sheet_name='이송재정리', index=False, startrow=0, startcol=0)
                cp_end_col = len(df_cp.columns)
            else:
                cp_end_col = 0 # 데이터가 없으면 시작 컬럼 0

            # 두 번째 데이터프레임 (CQ) - CP 데이터의 마지막 컬럼 + 1 에 저장
            if not df_cq.empty:
                df_cq.to_excel(writer, sheet_name='이송재정리', index=False, startrow=0, startcol=cp_end_col + 1)
                cq_end_col = cp_end_col + 1 + len(df_cq.columns)
            else:
                cq_end_col = cp_end_col + 1 # 데이터가 없으면 CP 마지막 컬럼 + 1 (컬럼이동)

            # 세 번째 데이터프레임 (CR) - CQ 데이터의 마지막 컬럼 + 1 에 저장
            if not df_cr.empty:
                df_cr.to_excel(writer, sheet_name='이송재정리', index=False, startrow=0, startcol=cq_end_col + 1)
                cr_end_col = cq_end_col + 1 + len(df_cr.columns)
            else:
                cr_end_col = cq_end_col + 1

            # 네 번째 데이터프레임 (CS) - CR 데이터의 마지막 컬럼 + 1 에 저장
            if not df_cs.empty:
                df_cs.to_excel(writer, sheet_name='이송재정리', index=False, startrow=0, startcol=cr_end_col + 1)
            
            print(f" - '이송재정리' 시트 데이터 병합 및 저장 완료.")


        print(f"\n모든 필터링된 데이터가 '{save_path}' (으)로 성공적으로 저장되었습니다.")

    except FileNotFoundError:
        print(f"오류: '{file_path}' 파일을 찾을 수 없습니다. 파일 경로를 확인해 주세요.")
    except ValueError as e:
        if "No sheet named 'ag-grid'" in str(e):
            print(f"오류: '{file_path}' 파일에 'ag-grid'라는 이름의 시트가 없습니다. 시트 이름을 확인해 주세요.")
        else:
            print(f"데이터를 처리하는 중 오류가 발생했습니다: {e}")
    except KeyError as e:
        print(f"오류: 필요한 컬럼이 데이터에 없습니다. '{e}' 컬럼이 있는지 확인해 주세요.")
    except Exception as e:
        print(f"예상치 못한 오류가 발생했습니다: {e}")

# 사용 예시 (Usage example)
# 여기에 실제 엑셀 파일 경로를 입력하세요.
# Replace 'your_excel_file.xlsx' with the actual path to your Excel file.
# 예를 들어, 현재 스크립트와 같은 폴더에 파일이 있다면 파일 이름만 적어도 됩니다.
excel_file = 'your_excel_file.xlsx' # 여기에 실제 파일 경로를 넣어주세요.
process_and_categorize_data(excel_file)