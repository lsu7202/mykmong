import os
import pandas as pd

def merge_excel_files(folder_path, output_file):
    files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]
    merged_data = []
    
    for file in files:
        file_path = os.path.join(folder_path, file)
        xls = pd.ExcelFile(file_path)
        
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            df.insert(0, 'Filename', file)  # 파일명 추가
            merged_data.append(df)
    
    final_df = pd.concat(merged_data, ignore_index=True)
    final_df.to_excel(output_file, index=False)
    
    print(f'Merged file saved as {output_file}')

# 사용 예시
folder_path = "/Users/iseung-ug/Desktop/보문산보물상점/야후제펜"  # 엑셀 파일들이 있는 폴더 경로
output_file = "merged_output.xlsx"  # 합친 파일 이름
merge_excel_files(folder_path, output_file)
