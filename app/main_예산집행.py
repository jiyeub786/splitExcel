import package.ExcelSplitProcessor as esp
import sys
def main():
    pj_nm = sys._getframe().f_code.co_filename
    print(pj_nm)

    path = "D:\\python\\project\\splitExcel\\app\\files\\02_예산집행"
    processor = esp (
        source_path=f"{path}/양식파일/작업본.xlsx" #작업본 경로
        ,template_path=f"{path}/양식파일/양식.xlsx" #양식 경로
        ,result_path=f"{path}/시도별" #분리파일 저장경로
        ,reulst_file_nm = "01_01_시설사업_예산집행 기준_데이터 정합성 검증 프로그램" #분리저장할 파일명칭
        ,result_file_date = '20240819_0900' #저장파일명칭 파일저장일시
    )

    # sido_list = ['서울','부산','대구','인천','광주','대전','울산','세종','경기','강원','충북','충남','전북','전남','경북','경남','제주']
    sido_list = ['강원']  # ,'부산','대구','인천','광주','대전','울산','세종','경기','강원','충북','충남','전북','전남','경북','경남','제주','교육']

    sheet_tasks = [
        # task0: 시트명,task1:  필터할 칼럼위치(a열 1 b열2 c열3 ...),task2:  복사할 범위,task3:  붙여넣을 범위
        ("학교목록", 2, "A2:G50000", "A2"),
        ("집행내역", 2, "A2:W500000", "A2"),
        ("통계분석_지역현황", 2, "A4:C500000", "A4")
    ]

    processor.process_sheets(sido_list, sheet_tasks)
    processor.quit()



#실행부
main()