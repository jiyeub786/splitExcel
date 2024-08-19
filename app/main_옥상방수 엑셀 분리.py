import package.ExcelSplitProcessor as esp
import sys
def main():
    pj_nm = sys._getframe().f_code.co_filename
    print(pj_nm)

    path = "D:\\python\\project\\splitExcel\\app\\files\\12_옥상방수쪼개기"
    processor = esp (
        source_path=f"{path}/양식파일/작업본.xlsx"
        ,template_path=f"{path}/양식파일/양식.xlsx"
        ,result_path=f"{path}/시도별"
        ,reulst_file_nm = "학교·기관 옥상방수 설치 운영현황"
        ,result_file_date = '20240819_0900'
    )

    # sido_list = ['서울','부산','대구','인천','광주','대전','울산','세종','경기','강원','충북','충남','전북','전남','경북','경남','제주']
    sido_list = ['강원']  # ,'부산','대구','인천','광주','대전','울산','세종','경기','강원','충북','충남','전북','전남','경북','경남','제주','교육']

    sheet_tasks = [
        # task0: 시트명,task1:  필터할 칼럼위치(a열 1 b열2 c열3 ...),task2:  복사할 범위,task3:  붙여넣을 범위
        ("학교·기관_운영내역", 2, "A1:BH20000", "A1"),
        ("옥상방수_시설내역", 2, "A1:W80000", "A1")
    ]

    processor.process_sheets(sido_list, sheet_tasks)
    processor.quit()



#실행부
main()