import win32com.client
import datetime


class ExcelSplitProcessor:
    def __init__(self, source_path, template_path, result_path,reulst_file_nm,result_file_date):
        self.excel = win32com.client.Dispatch("Excel.Application")
        self.excel.Visible = True
        self.excel.DisplayAlerts = False
        self.source_path = source_path
        self.template_path = template_path
        self.result_path = result_path
        self.reulst_file_nm = reulst_file_nm
        self.result_file_date = result_file_date


    def open_workbook(self, file_path):
        return self.excel.Workbooks.Open(file_path)

    def close_workbook(self, workbook, save_changes=False):
        workbook.Close(save_changes)

    def filter_column(self, worksheet, filter_index, filter_value):
        worksheet.Range("A1").AutoFilter(Field=filter_index, Criteria1=filter_value)

    def copy_sheet(self, soure_worksheet, target_worksheet, filter_index, filter_value, copy_range, paste_range):
        self.filter_column(soure_worksheet, filter_index, filter_value)
        filtered_range = soure_worksheet.AutoFilter.Range.SpecialCells(12)  # 필터된 셀 범위 가져오기
        filter_range = soure_worksheet.Range(copy_range)
        filter_range.Copy()
        target_worksheet.Range(paste_range).PasteSpecial()
        self.filter_column(target_worksheet, filter_index, filter_value)

    def set_init_workbook(self, workbook, zoom_level=75):
        sheet_cnt = workbook.Sheets.Count
        for sheet_num in range(sheet_cnt):
            workbook.Sheets(sheet_num + 1).Activate()
            workbook.Sheets(sheet_num + 1).Range("A1").Select() #초기 선택셀
            self.excel.ActiveWindow.Zoom = zoom_level #초기 확대값
        workbook.Sheets(1).Activate()

    def process_sheets(self, sido_list, sheet_tasks):
        soure_workbook = self.open_workbook(self.source_path)
        for i, sido in enumerate(sido_list):
            try:
                start_time = datetime.datetime.now()
                self.excel.Calculation = -4135  # 수식 끄기
                target_workbook = self.open_workbook(self.template_path)


                for task in sheet_tasks:
                    soure_worksheet1 = soure_workbook.Sheets(task[0])
                    target_worksheet1 = target_workbook.Sheets(task[0])
                    self.copy_sheet(soure_worksheet1, target_worksheet1,  task[1], sido, task[2], task[3])

                self.set_init_workbook(target_workbook)

                reulst_file_nm1 = f"{self.reulst_file_nm}_{str(i + 1).zfill(2)}_{sido}_{self.result_file_date}.xlsx"
                target_workbook.SaveAs(f"{self.result_path}/{reulst_file_nm1}")
                self.close_workbook(target_workbook, save_changes=True)

                self.excel.Calculation = -4105  # 수식 켜기
                elapsed_time = datetime.datetime.now() - start_time
                print(f"{sido} 처리 시간:", elapsed_time)
            except Exception as e:
                print(f"Error processing {sido}: {e}")
                self.close_workbook(target_workbook, save_changes=False)

        self.close_workbook(soure_workbook, save_changes=False)

    def quit(self):
        self.excel.Quit()