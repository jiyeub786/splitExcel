import os
import psutil
import pythoncom
import win32com.client
import datetime
import logging
from PySide6.QtCore import Signal
logger = logging.getLogger('MyLogger')
logger.setLevel(logging.DEBUG)
ch = logging.StreamHandler()  # 콘솔로 출력
ch.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(message)s')
ch.setFormatter(formatter)
logger.addHandler(ch)



class ExcelSplitProcessor:
    def __init__(self, source_path, template_path, result_path,reulst_file_nm,result_file_date
                    ,zoom_level1=120 #개요시트 zoom_level
                    ,zoom_level2=100 #나머지 시트 zoom_level
                    ,hide_guideline=True   #눈금선제거옵션 True 포함 False 제거
                 ):

        self.excel = ''
        self.source_path = source_path
        self.template_path = template_path
        self.result_path = result_path
        self.reulst_file_nm = reulst_file_nm
        self.result_file_date = result_file_date
        self.zoom_level1 = zoom_level1
        self.zoom_level2 = zoom_level2
        self.hide_guideline = hide_guideline



    def optimize_excel(self  ):
        """
        엑셀 성능 최적화를 위한 설정
        :param excel: Excel Application 객체
        """
        try:
            self.excel.ScreenUpdating = False  # 화면 업데이트 중지
            self.excel.DisplayAlerts = False  # 경고 메시지 비활성화
            self.excel.Calculation = -4135  # xlCalculationManual (수식 자동 계산 비활성화)
            #self.excel.EnableEvents = False  # 이벤트 비활성화
            self.excel.Visible = True  # 엑셀 창 숨김 (필요한 경우)
            #self.excel.DisplayStatusBar = False  # 상태 표시줄 비활성화
            #self.excel.AskToUpdateLinks = False  # 외부 링크 업데이트 요청 비활성화

            #엑셀에서 CPU 사용 개수를 조정하는 함수
            if True:
                self.excel.MultiThreadedCalculation.Enabled = True  # 멀티스레드 최대 활성화
                self.excel.MultiThreadedCalculation.ThreadCount = self.excel.MultiThreadedCalculation.ThreadCount
        except Exception as e:
            logger.info(f"optimize_excel(): {e}")

    def restore_excel(self  ):
        try:
            """
            엑셀 성능 설정을 원래대로 복구
            :param excel: Excel Application 객체
            """
            #self.excel.ScreenUpdating = True
            #self.excel.DisplayAlerts = True
            self.excel.Calculation = -4105  # xlCalculationAutomatic (수식 자동 계산 활성화)
            #self.excel.EnableEvents = True
            #self.excel.Visible = True  # 다시 표시
        except Exception as e:
            logger.info(f"restore_excel(): {e}")
    def open_workbook(self, file_path):
        return self.excel.Workbooks.Open(file_path)

    def close_workbook(self, workbook, save_changes=False):
        try:
            workbook.Close(save_changes)
        except Exception as e:
            logger.info(f"close_workbook(): {e}")

    def filter_column(self, worksheet, filter_index, filter_value):
        try:
            worksheet.Range("A1").AutoFilter(Field=filter_index, Criteria1=filter_value)
        except Exception as e:
            logger.info(f"filter_column(): {e}")

    def clear_autofilter(self,worksheet):
        try:
            if worksheet.AutoFilterMode:
                worksheet.AutoFilterMode = False ##"""워크시트에서 AutoFilter를 해제하는 함수"""
        except Exception as e:
            logger.info(f"clear_autofilter(): {e}")
    def delete_rows_after_last_data(self, worksheet , column="A"):
        try:
            last_row = worksheet.Cells(worksheet.Rows.Count, column).End(-4162).Row  # xlUp(-4162) 사용
            total_rows = worksheet.Rows.Count  # 엑셀의 최대 행 수
            self.clear_autofilter(worksheet)
            if last_row < total_rows:  # 마지막 행이 엑셀의 끝이 아닐 경우만 실행
                worksheet.Rows(f"{last_row + 1}:{total_rows}").Delete()
        except Exception as e:
            logger.info(f"delete_rows_after_last_data(): {e}")


    def copyAndPaste_sheet(self, soure_worksheet, target_worksheet, filter_index, filter_value, copy_range, paste_range):

        try:
            self.filter_column(soure_worksheet, filter_index, filter_value)
            #filtered_range = soure_worksheet.AutoFilter.Range.SpecialCells(12)  # 필터된 셀 범위 가져오기
            filter_range = soure_worksheet.Range(copy_range)
            filter_range.Copy()
            target_worksheet.Range(paste_range).PasteSpecial()

            #self.filter_column(target_worksheet, filter_index, filter_value)
            return 1
        except Exception as e:
            logger.info(f"copyAndPaste_sheet(): {e}")
            return 0

    def hide_gridlines(self):
        try:
            self.excel.ActiveWindow.DisplayGridlines = self.hide_guideline
            #print(self.hide_guideline)
        except Exception as e:
            logger.info(f"hide_gridlines(): {e}")

    def set_init_workbook(self, workbook ):
        try:
            sheet_cnt = workbook.Sheets.Count
            for sheet_num in range(sheet_cnt):
                self.hide_gridlines()
                workbook.Sheets(sheet_num + 1).Activate()
                workbook.Sheets(sheet_num + 1).Range("A1").Select() #초기 선택셀
                if sheet_num + 1 == 1:
                    self.excel.ActiveWindow.Zoom = self.zoom_level1 #초기 확대값
                else :
                    self.excel.ActiveWindow.Zoom = self.zoom_level2  # 초기 확대값
            workbook.Sheets(1).Activate()
        except Exception as e:
            logger.info(f"set_init_workbook(): {e}")


    def process_sheets(self, sido_list, sheet_tasks):

        try:
            logger.info("엑셀 분리작업을 시작합니다")
            #pythoncom.CoInitialize()
            self.excel = win32com.client.Dispatch("Excel.Application")

            soure_workbook = self.open_workbook(self.source_path)
            self.optimize_excel() ##성능최적화 켜기

            for i, sido in enumerate(sido_list):
                start_time = datetime.datetime.now()
                target_workbook = self.open_workbook(self.template_path)
                self.optimize_excel() ##성능최적화 켜기

                for task in sheet_tasks:
                    soure_worksheet1 = soure_workbook.Sheets(task[0])
                    target_worksheet1 = target_workbook.Sheets(task[0])

                    if 1 == self.copyAndPaste_sheet(soure_worksheet1, target_worksheet1,  task[1], sido, task[2], task[3]):
                        self.delete_rows_after_last_data(target_worksheet1)


                self.set_init_workbook(target_workbook)
                self.restore_excel() ##성능최적화 끄기

                reulst_file_nm1 = f"{self.reulst_file_nm}_{str(i + 1).zfill(2)}_{sido}_{self.result_file_date}.xlsx"
                target_workbook.SaveAs(f"{self.result_path}/{reulst_file_nm1}")
                self.close_workbook(target_workbook, save_changes=True)

                elapsed_time = datetime.datetime.now() - start_time
                logger.info(f"{sido} 처리 시간: {elapsed_time}")

            logger.info(f"모든작업 완료")
            self.close_workbook(soure_workbook, save_changes=False)

        except Exception as e:
            logger.error(f"process_sheets() Error processing {sido}: {e}")

        finally:

            # EXCEL.EXE 프로세스 정리
            if self.excel is not None:
                for proc in psutil.process_iter():
                    if proc.name().lower() == "excel.exe":
                        proc.kill()  # 프로세스 강제 종료
            #pythoncom.CoUninitialize()  # COM 해제


    def quit(self):

        # EXCEL.EXE 프로세스 정리
        for proc in psutil.process_iter():
            if proc.name().lower() == "excel.exe":
                print("프로세스 강제종료")
                print("프로세스 강제종료2")
                proc.kill()  # 프로세스 강제 종료
                print("프로세스 강제종료3")

