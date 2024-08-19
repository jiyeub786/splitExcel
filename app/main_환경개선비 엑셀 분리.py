import package.ExcelSplitProcessor as esp
import sys
def main():
    pj_nm = sys._getframe().f_code.co_filename
    print(pj_nm)

    path = "D:\\python\\project\\splitExcel\\app\\files\\22_환경개선비"
    processor = esp (
        source_path=f"{path}/양식파일/작업본.xlsx" #작업본 경로
        ,template_path=f"{path}/양식파일/양식.xlsx" #양식 경로
        ,result_path=f"{path}/시도별" #분리파일 저장경로
        ,reulst_file_nm = "09_2025년 환경개선 대상시설 물량 산출" #분리저장할 파일명칭
        ,result_file_date = '20240819_0900' #저장파일명칭 파일저장일시
    )

    # sido_list = ['서울','부산','대구','인천','광주','대전','울산','세종','경기','강원','충북','충남','전북','전남','경북','경남','제주']
    sido_list = ['강원']  # ,'부산','대구','인천','광주','대전','울산','세종','경기','강원','충북','충남','전북','전남','경북','경남','제주','교육']

    sheet_tasks = [
        # task0: 시트명,task1:  필터할 칼럼위치(a열 1 b열2 c열3 ...),task2:  복사할 범위,task3:  붙여넣을 범위
        ("건물_공간_시설현황", 2, "A3:AL50000", "A3"),
        ("대상시설_목록", 2, "A3:AA500000", "A3"),
        ("학교단위_목록", 2, "A3:K20000", "A3"),
        ("학교단위_목록", 2, "P3:S20000", "P3")
    ]

    processor.process_sheets(sido_list, sheet_tasks)
    processor.quit()



#실행부
main()