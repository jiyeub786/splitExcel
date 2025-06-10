import gc
import os
import time
import pythoncom
from win32com.client import CDispatch
import win32com.client
import datetime
import logging
os.environ["PYTHONUNBUFFERED"] = "1"
logger = logging.getLogger('MyLogger')
logger.setLevel(logging.DEBUG)
ch = logging.StreamHandler()  # 콘솔로 출력
ch.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(message)s')
ch.setFormatter(formatter)
logger.addHandler(ch)

def read_flag():
    try:
        with open("flag.txt", "r") as f:
            return f.read().strip()
    except:
        return "0"  # 플래그가 없으면 종료

class ExcelSplitProcessor:
    def __init__(self, source_path, template_path, result_path,reulst_file_nm,result_file_date
                    ,option_zoom_level1=120 #개요시트 zoom_level
                    ,option_zoom_level2=100 #나머지 시트 zoom_level
                    ,option_hide_guidelineYN= "Y"   #눈금선제거옵션 Y켜기 N 끄기
                    ,option_formula_removePathYN ="Y"
                 ):

        self.excel = ''
        self.source_path = source_path
        self.template_path = template_path
        self.result_path = result_path
        self.reulst_file_nm = reulst_file_nm
        self.result_file_date = result_file_date
        self.option_zoom_level1 = int(option_zoom_level1)
        self.option_zoom_level2 = int(option_zoom_level2)
        self.option_hide_guidelineYN = option_hide_guidelineYN
        self.option_formula_removePathYN = option_formula_removePathYN
        self.target_workbook = ''
        self.source_workbook = ''



    def optimize_excel(self  ):
        """
        엑셀 성능 최적화를 위한 설정
        :param excel: Excel Application 객체
        """
        try:
            #self.excel.ScreenUpdating = False  # 화면 업데이트 중지
            if self.excel.DisplayAlerts != False:
                self.excel.DisplayAlerts = False  # 경고 메시지 비활성화

            if self.excel.Calculation != -4135:
                self.excel.Calculation = -4135  # xlCalculationAutomatic (수식 자동 계산 비활성화)

            if self.excel.EnableEvents != False:
                self.excel.EnableEvents = False  # 이벤트 비활성화

            if self.excel.DisplayStatusBar != False:
                self.excel.DisplayStatusBar = False  # 상태 표시줄 비활성화

            if self.excel.AskToUpdateLinks != False:
                self.excel.AskToUpdateLinks = False  # 외부 링크 업데이트 요청 비활성화

            # #엑셀에서 CPU 사용 개수를 조정하는 함수
            # if True:
            #     self.excel.MultiThreadedCalculation.Enabled = True  # 멀티스레드 최대 활성화
            #     self.excel.MultiThreadedCalculation.ThreadCount = self.excel.MultiThreadedCalculation.ThreadCount
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
            if self.excel.Calculation != -4105:
                self.excel.Calculation = -4105  # xlCalculationAutomatic (수식 자동 계산 활성화)
            #self.excel.EnableEvents = True
            #self.excel.Visible = True  # 다시 표시
        except Exception as e:
            logger.info(f"restore_excel(): {e}")
    def open_workbook(self, file_path):
        return self.excel.Workbooks.Open(file_path)

    def close_workbook(self, workbook, save_changes=False):
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=save_changes)
        except Exception as e:
            logger.info(f"close_workbook(): {e}")

    def quit_excel(self, app):
        try:
            if app is not None:
                app.Quit()
        except Exception as e:
            logger.warning(f"quit_excel() error: {e}")
        finally:
            # COM 객체 해제 및 가비지 컬렉션
            del app
            gc.collect()


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


    def remove_pathText(self,txt):
        fileName = self.source_path
        path = os.path.dirname(fileName).replace("/", "\\") + '\\'
        filename = "[" + os.path.basename(fileName).replace("/", "\\") + "]"
        return txt.replace(path, "").replace(filename, "")

    # def remove_external_references_from_worksheet(self, worksheet):
    #     print("remove")
    #     used_range = worksheet.UsedRange
    #     for row in used_range.Rows:
    #         for cell in row.Cells:
    #             try:
    #                 if cell.HasFormula:
    #                     formula = cell.Formula
    #                     # 외부 참조 있는 경우
    #                     if "[" in formula and "]" in formula:
    #                         new_formula = self.remove_pathText(formula)
    #                         cell.Formula = new_formula
    #             except Exception as e:
    #                 logger.info(f"remove_external_references_from_worksheet(): {e}")

    def remove_external_references_from_worksheet(self, worksheet):
        print("remove")
        try:
            used_range = worksheet.UsedRange
            formulas = used_range.Formula  # 2D array (row x col)

            row_count = used_range.Rows.Count
            col_count = used_range.Columns.Count

            # 변경된 formula 저장용 배열
            updated_formulas = [[None for _ in range(col_count)] for _ in range(row_count)]

            changed = False

            for i in range(row_count):
                for j in range(col_count):
                    cell_formula = formulas[i][j]
                    if isinstance(cell_formula, str) and "[" in cell_formula and "]" in cell_formula:
                        new_formula = self.remove_pathText(cell_formula)
                        updated_formulas[i][j] = new_formula
                        changed = True
                    else:
                        updated_formulas[i][j] = cell_formula  # 그대로 유지

            if changed:
                used_range.Formula = updated_formulas  # 한 번에 반영

        except Exception as e:
            logger.info(f"remove_external_references_from_worksheet(): {e}")


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
            self.excel.ActiveWindow.DisplayGridlines = False
            #print(self.hide_guideline)
        except Exception as e:
            logger.info(f"hide_gridlines(): {e}")

    def set_init_workbook(self, workbook):
        try:
            sheet_cnt = workbook.Sheets.Count
            for sheet_num in range(sheet_cnt):
                sheet = workbook.Sheets(sheet_num + 1)

                # 숨겨진 시트는 건너뜀
                if not sheet.Visible:
                    continue

                # 눈금선 제거 옵션 (Y: 켜짐, N: 꺼짐)
                if self.option_hide_guidelineYN == "N":
                    self.hide_gridlines()

                sheet.Activate()
                sheet.Range("A1").Select()  # 초기 선택 셀
                self.excel.ActiveWindow.ScrollRow = 1  # 화면의 최상단으로 이동
                self.excel.ActiveWindow.ScrollColumn = 1  # 화면의 최좌측으로 이동

                # 확대 비율 적용
                if sheet_num + 1 == 1:
                    if self.option_zoom_level1 != 0:
                        self.excel.ActiveWindow.Zoom = self.option_zoom_level1
                else:
                    if self.option_zoom_level2 != 0:
                        self.excel.ActiveWindow.Zoom = self.option_zoom_level2

            # 첫 번째 보이는 시트로 이동
            for sheet_num in range(sheet_cnt):
                if workbook.Sheets(sheet_num + 1).Visible:
                    workbook.Sheets(sheet_num + 1).Activate()
                    break

        except Exception as e:
            logger.info(f"set_init_workbook(): {e}")

    def process_sheets(self, tgt_list, sheet_tasks):
        print("S1")
        self.target_workbook =None
        self.soure_workbook =None
        #self.excel =None

        try:
            logger.info("엑셀 분리작업을 시작합니다")
            #pythoncom.CoInitialize()
            print("S2")

            if self.excel is None or not hasattr(self.excel, "Workbooks"):
                self.excel = win32com.client.Dispatch("Excel.Application")

            print("S3")
            if self.soure_workbook is None or not hasattr(self.soure_workbook, "Close"):
                self.soure_workbook = self.open_workbook(self.source_path)

            print("S4")

            self.optimize_excel() ##성능최적화 켜기
            total_start_time = datetime.datetime.now()

            for i, tgt in enumerate(tgt_list):
                print("1")
                self.optimize_excel() ##성능최적화 켜기
                each_start_time = datetime.datetime.now()
                self.target_workbook = self.open_workbook(self.template_path)
                print("2")

                print("3")

                for task in sheet_tasks:

                    print("l-1")
                    soure_worksheet1 = self.soure_workbook.Sheets(task[0])
                    print("l-2")
                    target_worksheet1 = self.target_workbook.Sheets(task[0])

                    # 숨겨진 시트는 건너뜀
                    if not soure_worksheet1.Visible or not target_worksheet1.Visible:
                        logger.info(f"숨겨진 시트는 작업을 할 수 없어 건너뜁니다. 시트명-{task[0]}")
                        continue

                    if read_flag() == "1":
                        if 1 == self.copyAndPaste_sheet(soure_worksheet1, target_worksheet1,  task[1], tgt, task[2], task[3]):
                            self.delete_rows_after_last_data(target_worksheet1)
                            if self.option_formula_removePathYN == "Y": #계산식에서 경로제거
                                self.remove_external_references_from_worksheet(target_worksheet1)


                    if read_flag() == "0":
                        raise RuntimeError("종료 플래그 감지")



                print("l-3")
                self.set_init_workbook(self.target_workbook)
                print("l-4")
                self.restore_excel() ##성능최적화 끄기
                print("l-5")

                reulst_file_nm1 = f"{self.reulst_file_nm}_{str(i + 1).zfill(2)}_{tgt}_{self.result_file_date}.xlsx"
                print("l-6")
                self.target_workbook.SaveAs(f"{self.result_path}/{reulst_file_nm1}")
                print("l-7")
                #time.sleep(0.5)
                self.close_workbook(self.target_workbook, save_changes=False)
                print("l-8")
                self.target_workbook =None


                elapsed_time = datetime.datetime.now() - each_start_time
                logger.info(f"{tgt} 완료. 처리 시간: {int(elapsed_time.total_seconds())}초")


            print("4")

            self.close_workbook(self.soure_workbook, save_changes=False)
            print("5")
            self.soure_workbook = None

            elapsed_time = datetime.datetime.now() - total_start_time
            logger.info(f"모든작업 완료. 총 처리 시간: {int(elapsed_time.total_seconds())}초")

            #self.close_workbook(soure_workbook, save_changes=False)
            #self.quit_excel(self.excel)




        except Exception as e:
            logger.error(f"process_sheets() Error processing: {e}")
            self.target_workbook = None

            self.close_workbook(self.soure_workbook, save_changes=False)
            self.soure_workbook = None
            #self.quit_excel(self.excel)
            pythoncom.CoUninitialize()



