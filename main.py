import ctypes
import sys
import threading
from warnings import catch_warnings

import win32com
from PySide6.QtCore import QCoreApplication, QThread, Signal
from PySide6.QtWidgets import QMainWindow, QApplication, QFileDialog, QTableWidgetItem, QPlainTextEdit, QVBoxLayout, \
    QDialog
from myUi import Ui_MainWindow  # 만약 파이썬 파일이름이 component.py 라면 from component .. 로 진행

from logUi import Ui_Dialog  # 만약 파이썬 파일이름이 component.py 라면 from component .. 로 진행

import json

import package.ExcelSplitProcessor as esp

import logging

# 로깅 설정
logger = logging.getLogger('MyLogger')
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(message)s')


msg01= "원본파일을 선택해주세요"
msg02= "양식파일을 선택해주세요"
msg03= "결과값을 저장할 폴더를 선택해주세요"



class WorkerThread(QThread):
    def __init__(self):
        super().__init__()
        self.file01 = ''
        self.file02 = ''
        self.folder01 = ''
        self.saveNm01 = ''
        self.saveNm02 = ''
        self.workList01 = ''
        self.workList02 = ''
        self.testYn = ''
        self.processor = ''




    def setValues(self, file01 ,file02 ,folder01,saveNm01,saveNm02,workList01,workList02,testYn):
        self.file01 = file01
        self.file02 = file02
        self.folder01 = folder01
        self.saveNm01 = saveNm01
        self.saveNm02 = saveNm02
        self.workList01 =workList01
        self.workList02 = workList02
        self.testYn = testYn

    def stop(self):
        self.terminate()
        if self.processor != '':
            print("whdf")
            self.processor.quit()


    def run(self):

        try:
            self.processor  = esp(
                ## reulst_file_nm	분할파일명칭
                ## result_file_date	파일저장일시
                ## zoom_level1	개요시트 확대 수치, 기본값 120%
                ## zoom_level2	나머지시트 확대 수치, 기본값 75%
                ## hide_guideline	엑셀 눈금선 끄기(True), 켜기(False)
                source_path=self.file01.replace('/', '\\')  # 작업본 경로
                , template_path=self.file02.replace('/', '\\')  # 양식 경로
                , result_path=self.folder01.replace('/', '\\')  # 분리파일 저장경로
                , reulst_file_nm=self.saveNm01  # 분리저장할 파일명칭
                , result_file_date=self.saveNm02  # 저장파일명칭 파일저장일시
                , zoom_level1=145  # 개요시트 zoom level
                , zoom_level2=75  # 나머지시트 zoom level
                , hide_guideline=False  # 눈금선 제거 옵션
            )


            logger.info(f"------설정값------------------------")
            logger.info(f"원본파일:{self.file01}")
            logger.info(f"양식파일:{self.file02}")
            logger.info(f"저장폴더:{self.folder01}")
            logger.info(f"저장형식:{self.saveNm01}_**_{self.saveNm02}.xlsx")

            logger.info(f"작업대상1:{self.workList01}")
            logger.info(f"작업정의2:{self.workList02}")
            logger.info(f"----------------------------------")


            ## STEP - 02
            if self.testYn == 'Y':
                # 테스트용으로 리스트중 맨앞만 실행
                work = self.workList02[0:1]

            else:
                work = self.workList02

            split_list = work

            ## STEP - 03
            sheet_tasks = self.workList01

            self.processor.process_sheets(split_list, sheet_tasks)
        except Exception as e:
            logger.info(f"에러발생\n {e}")
            logger.info(f"작업중단")

        finally:
            self.processor.quit()
            self.processor.quit_excel()


class LogWindow(QDialog):
    def __init__(self):
        super().__init__()
        self.ui = Ui_Dialog()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.event_bttnExit )
        # 로깅 포맷터 설정
        self.formatter = logging.Formatter('%(asctime)s - %(message)s')
        
        log_handler = logging.StreamHandler()
        log_handler.setLevel(logging.DEBUG)
        log_handler.setFormatter(formatter)

        # 로그를 text_edit에 출력하도록 설정
        log_handler.emit = self.emit_log

        logger.addHandler(log_handler)

        # 로그 시작
        logger.info("로그창이 시작되었습니다.")

    def emit_log(self, record):
        """로그를 QPlainTextEdit에 출력하는 함수"""
        log_message = self.formatter.format(record)
        self.ui.plainTextEdit.appendPlainText(log_message)

    def event_bttnExit(self):
        print("종료")
        self.close()


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.worker_thread = None
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.log_window = LogWindow()

        self.setInitUiText()

        self.configFileNm = "config.json"
        self.file01 = ''
        self.file02 = ''
        self.folder01 = ''

        self.saveNm01 = ''
        self.saveNm02 = ''

        self.workList01 =[]
        self.workList02 = []
        self.workList03 = ['','','']

        self.threadFlag = 0



     #버튼 이벤트
        self.ui.bttnLoad01.clicked.connect(self.event_bttnLoad01 )
        self.ui.bttnLoad02.clicked.connect(self.event_bttnLoad02 )
        self.ui.bttnLoad03.clicked.connect(self.event_bttnLoad03 )
        self.ui.bttnSaveConfig.clicked.connect(self.saveConfig )
        self.ui.bttnLoadConfig.clicked.connect(self.loadConfig )


        self.ui.bttnShowLog.clicked.connect(self.event_showLogUi )
        self.ui.bttnForceStop.clicked.connect(self.event_bttnForceStop )

        self.ui.bttnStartWork.clicked.connect(self.event_bttnStartWork)
        self.ui.bttnStartTest.clicked.connect(self.event_bttnStartTest)

        self.ui.bttnCreateRow01.clicked.connect(self.event_table01_InsertRow )
        self.ui.bttnCreateRow02.clicked.connect(self.event_table02_InsertRow )


        #self.ui.plainTextEdit01.(self.event_bttnLoad02 )


        self.ui.bttnExit.clicked.connect(self.event_bttnExit )
     #변수 설정

        self.list1 = ['1a', '2b', '3c', '4d']
        self.list2 = ['1ac', '2bd', '3ce', '4df']




    def saveConfig(self):
        self.getTextValues()



        config_data = {
            "file01": self.file01,
            "file02": self.file02,
            "folder01": self.folder01,
            "saveNm01": self.saveNm01,
            "saveNm02": self.saveNm02,
            "workList01": self.workList01,
            "workList02": self.workList02 
        }

        fsaveNm = QFileDialog.getSaveFileName(self, "파일 저장")

        with open(fsaveNm[0], "w", encoding="utf-8") as file:
            json.dump(config_data, file, indent=4, ensure_ascii=False)



        print(config_data)

    def loadConfig(self):
        loadFileNm = QFileDialog.getOpenFileName(self, "파일 열기" )
        """JSON 파일에서 설정을 불러와 딕셔너리에 저장"""
        try:
            with open(loadFileNm[0], "r", encoding="utf-8") as file:
                config_data = json.load(file)  # JSON 데이터를 딕셔너리로 변환

                print(config_data)
            # 클래스 변수에 JSON 값 자동 할당
            for key, value in config_data.items():
                if hasattr(self, key):  # 해당 키가 클래스 변수로 존재하는 경우
                    setattr(self, key, value)

            self.refeshUiText()

        except (FileNotFoundError, json.JSONDecodeError):
            print("설정 파일이 없거나 형식이 잘못되었습니다.")

    # 설정값 로드 후 변수 재할당
    def refeshUiText(self):
        self.ui.plainTextEdit01.setPlainText(self.file01 )
        self.ui.plainTextEdit02.setPlainText(self.file02 )
        self.ui.plainTextEdit03.setPlainText(self.folder01 )
        self.ui.plainTextEdit04.setPlainText(self.saveNm01 )
        self.ui.plainTextEdit05.setPlainText(self.saveNm02 )

        self.table01_setData()
        self.table02_setData()



    def setInitUiText(self):
        self.ui.plainTextEdit01.setPlainText(QCoreApplication.translate("MainWindow",msg01, None))
        self.ui.plainTextEdit02.setPlainText(QCoreApplication.translate("MainWindow",msg02, None))
        self.ui.plainTextEdit03.setPlainText(QCoreApplication.translate("MainWindow",msg03, None))

    def getTextValues(self):
        self.file01 = self.ui.plainTextEdit01.toPlainText()
        self.file02 = self.ui.plainTextEdit02.toPlainText()
        self.folder01 = self.ui.plainTextEdit03.toPlainText()
 


        self.saveNm01 = self.ui.plainTextEdit04.toPlainText()
        self.saveNm02 = self.ui.plainTextEdit05.toPlainText()
        self.table01_saveData()
        self.table02_saveData()



    def event_bttnExit(self):
        print("종료")
        QCoreApplication.instance().quit()

    def event_bttnLoad01(self):
        f =QFileDialog.getOpenFileName(self,'','','Excel(*.xlsx *xls)')
        if f[0]:
            self.ui.plainTextEdit01.setPlainText(f[0])
            self.file01 = f[0]
        else:
            self.ui.plainTextEdit01.setPlainText(msg01)
            print(msg01)


    def event_bttnLoad02(self):
        f =QFileDialog.getOpenFileName(self,'','','Excel(*.xlsx *xls)')
        if f[0]:
            self.ui.plainTextEdit02.setPlainText(f[0])
            self.file02 = f[0]
            print(f)
        else:
            self.ui.plainTextEdit02.setPlainText(msg02)
            print(msg02)


    def event_bttnLoad03(self):
        fd =QFileDialog.getExistingDirectory(self,'폴더선택','')
        if fd:
            self.ui.plainTextEdit03.setPlainText(fd)
            self.folder01 = fd
            print(fd)
        else:
            self.ui.plainTextEdit03.setPlainText(msg03)
            print(msg03)

    def event_table01_InsertRow(self):
        row_count = self.ui.table01.rowCount()
        self.ui.table01.insertRow(row_count)

    def event_table02_InsertRow(self):
        row_count = self.ui.table02.rowCount()
        self.ui.table02.insertRow(row_count)


    def table01_saveData(self):
        """테이블 데이터를 저장"""
        data = []
        row_count = self.ui.table01.rowCount()
        col_count = self.ui.table01.columnCount()

        for row in range(row_count):
            row_data = []
            for col in range(col_count):
                item = self.ui.table01.item(row, col)
                row_data.append(item.text() if item else "")  # 빈 셀은 빈 문자열로 처리
            data.append(row_data)
        self.workList01 = data.copy()

        print("저장된 데이터:", data)  # 출력 (실제로는 파일 저장이나 DB 저장 가능)


    def table02_saveData(self):
        """테이블 데이터를 저장"""
        data = []
        row_count = self.ui.table02.rowCount()
        col_count = self.ui.table02.columnCount()

        for row in range(row_count):
            for col in range(col_count):
                item = self.ui.table02.item(row, col)
                data.append(item.text())

        self.workList02 = data.copy()


    def table01_setData(self):
        """테이블에 초기 데이터를 설정하는 함수"""
        initial_data = self.workList01

        self.ui.table01.setRowCount(len(initial_data))  # 행 개수 설정

        for row, rowData in enumerate(initial_data):
            for col, value in enumerate(rowData):
                self.ui.table01.setItem(row, col, QTableWidgetItem(value))

    def table02_setData(self):
        """테이블에 초기 데이터를 설정하는 함수"""
        initial_data = self.workList02

        self.ui.table02.setRowCount(len(initial_data))  # 행 개수 설정

        for row, rowData in enumerate(initial_data):
            self.ui.table02.setItem(row,0, QTableWidgetItem(rowData))





    def event_bttnStartWork(self):
        try:
            self.threadFlag = 1
            logger.info("작업을 시작합니다")
            self.getTextValues()
            self.worker_thread = WorkerThread()
            self.worker_thread.setValues( self.file01 ,self.file02 ,self.folder01 ,self.saveNm01 ,self.saveNm02,self.workList01,self.workList02 ,"N")


            self.worker_thread.start()
            self.event_showLogUi()  # 로그 UI 표시
        except Exception as e:
            logger.info(f"작업중 에러가 발생했습니다\n{e}")

    def event_bttnStartTest(self):
        try:
            self.threadFlag  = 1
            logger.info("작업 테스트를 시작합니다")
            self.getTextValues()
            self.worker_thread = WorkerThread()
            self.worker_thread.setValues(self.file01, self.file02, self.folder01, self.saveNm01, self.saveNm02,self.workList01, self.workList02 ,"Y")


            self.worker_thread.start()
            self.event_showLogUi()  # 로그 UI 표시
        except Exception as e:
            logger.info(f"작업중 에러가 발생했습니다\n{e}")


    def event_bttnForceStop(self):
        try:
            if self.threadFlag == 1:
                logger.info("작업을 강제로 종료합니다")
                self.worker_thread.stop()
                logger.info("작업이 강제 종료되었습니다")
            else :
                logger.info("실행중인 작업이 없습니다")
        except Exception as e:
            logger.info(f"실행중인 작업이 없거나, 작업강제 종료중 에러가 발생했습니다\n{e}")

    def event_showLogUi(self):
        self.log_window.show()


if __name__ == "__main__":
    app = QApplication()
    window = MainWindow()
    window.show()
    sys.exit(app.exec())