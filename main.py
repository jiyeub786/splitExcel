
import os
import sys
import psutil
from PySide6.QtCore import QCoreApplication, QThread, Signal, QObject
from PySide6.QtGui import QTextCursor
from PySide6.QtWidgets import QMainWindow, QApplication, QFileDialog, QTableWidgetItem,  QDialog
from myUi import Ui_MainWindow  # 만약 파이썬 파일이름이 component.py 라면 from component .. 로 진행
from usageUi import Ui_Dialog  as usageUi_Dialog# 만약 파이썬 파일이름이 component.py 라면 from component .. 로 진행
from logUi import Ui_Dialog  as logUi_Dialog# 만약 파이썬 파일이름이 component.py 라면 from component .. 로 진행

import json
import package.ExcelSplitProcessor as esp
import logging

def writeFlag(value):
    # 플래그 파일 생성
    with open("flag.txt", "w") as f:
        f.write(value)  # 1이면 계속 실행, 0이면 종료value


os.environ["PYTHONUNBUFFERED"] = "1"

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
        self.workList03 = ''
        self.testYn = ''
        self.processor = ''




    def setValues(self, file01 ,file02 ,folder01,saveNm01,saveNm02,workList01,workList02,workList03,testYn):
        self.file01 = file01
        self.file02 = file02
        self.folder01 = folder01
        self.saveNm01 = saveNm01
        self.saveNm02 = saveNm02
        self.workList01 =workList01
        self.workList02 = workList02
        self.workList03 = workList03
        self.testYn = testYn

    def stop(self):
        try:
            writeFlag("0")
        except Exception as e:
            logger.info(f"stop()\n {e}")



    def run(self):

        writeFlag("1")
        try:
            self.processor  = esp(
                source_path=self.file01.replace('/', '\\')  # 작업본 경로
                , template_path=self.file02.replace('/', '\\')  # 양식 경로
                , result_path=self.folder01.replace('/', '\\')  # 분리파일 저장경로
                , reulst_file_nm=self.saveNm01  # 분리저장할 파일명칭
                , result_file_date=self.saveNm02  # 저장파일명칭 파일저장일시
                , option_zoom_level1= self.workList03[0]  # 개요시트 zoom level
                , option_zoom_level2= self.workList03[1]  # 나머지시트 zoom level
                , option_hide_guidelineYN= self.workList03[2]  # 눈금선 제거 옵션
                , option_formula_removePathYN= self.workList03[3]  # 눈금선 제거 옵션
            )


            logger.info(f"------설정값------------------------")

            logger.info(f"원본파일:{self.file01}")
            logger.info(f"양식파일:{self.file02}")
            logger.info(f"저장폴더:{self.folder01}")
            logger.info(f"저장형식:{self.saveNm01}_**_{self.saveNm02}.xlsx")

            logger.info(f"작업대상1:{self.workList01}")
            logger.info(f"작업정의2:{self.workList02}")
            logger.info(f"작업정의3:{self.workList03}")

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




class UsageWindow(QDialog):
    def __init__(self):
        super().__init__()
        self.ui = usageUi_Dialog()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.event_bttnExit)

    def event_bttnExit(self):
        print("종료")
        self.close()

class LogSignal(QObject):
    log_signal = Signal(str)  # 로그 텍스트를 전달하는 시그널

class LogWindow(QDialog):
    def __init__(self):
        super().__init__()
        self.ui = logUi_Dialog()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.event_bttnExit )
        # 로깅 포맷터 설정
        # 시그널 객체 생성
        self.log_signal = LogSignal()
        self.log_signal.log_signal.connect(self.append_log)

        # 로깅 포맷터 설정
        self.formatter = logging.Formatter('%(asctime)s - %(message)s')

        # 커스텀 핸들러 생성
        log_handler = logging.StreamHandler()
        log_handler.setLevel(logging.DEBUG)
        log_handler.setFormatter(self.formatter)
        log_handler.emit = self.emit_log  # emit 메서드 오버라이드

        logger.addHandler(log_handler)

        logger.info("로그창이 시작되었습니다.")

    def emit_log(self, record: logging.LogRecord):
        """StreamHandler의 emit 메서드를 오버라이드해서 시그널로 전달"""
        try:
            message = self.formatter.format(record)
            self.log_signal.log_signal.emit(message)  # 시그널로 전달
        except Exception as e:
            print(f"emit_log 오류: {e}")

    def append_log(self, message: str):
        """로그를 QPlainTextEdit에 출력"""
        try:
            self.ui.plainTextEdit.appendPlainText(message)
            self.ui.plainTextEdit.moveCursor(QTextCursor.End)
            self.ui.plainTextEdit.ensureCursorVisible()
        except Exception as e:
            logger.info(f"append_log: {e}")

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
        self.usage_window = UsageWindow()

        self.setInitUiText()

        self.configFileNm = "config.json"
        self.file01 = ''
        self.file02 = ''
        self.folder01 = ''

        self.saveNm01 = ''
        self.saveNm02 = ''

        self.workList01 =[]
        self.workList02 = []
        self.workList03 = []

        self.threadFlag = 0



     #버튼 이벤트
        self.ui.bttnLoad01.clicked.connect(self.event_bttnLoad01 )
        self.ui.bttnLoad02.clicked.connect(self.event_bttnLoad02 )
        self.ui.bttnLoad03.clicked.connect(self.event_bttnLoad03 )
        self.ui.bttnSaveConfig.clicked.connect(self.saveConfig )
        self.ui.bttnLoadConfig.clicked.connect(self.loadConfig )


        self.ui.bttnShowLog.clicked.connect(self.event_showLogUi )
        self.ui.bttnShowUsage.clicked.connect(self.event_showUsageUi )
        self.ui.bttnForceStop.clicked.connect(self.event_bttnForceStop )

        self.ui.bttnStartWork.clicked.connect(self.event_bttnStartWork)
        self.ui.bttnStartTest.clicked.connect(self.event_bttnStartTest)

        self.ui.bttnCreateRow01.clicked.connect(self.event_table01_InsertRow )
        self.ui.bttnCreateRow02.clicked.connect(self.event_table02_InsertRow )


        #self.명령어.plainTextEdit01.(self.event_bttnLoad02 )


        self.ui.bttnExit.clicked.connect(self.event_bttnExit )
     #변수 설정

        self.list1 = ['1a', '2b', '3c', '4d']
        self.list2 = ['1ac', '2bd', '3ce', '4df']




    def saveConfig(self):
        try:
            self.getTextValues()



            config_data = {
                "file01": self.file01,
                "file02": self.file02,
                "folder01": self.folder01,
                "saveNm01": self.saveNm01,
                "saveNm02": self.saveNm02,
                "workList01": self.workList01,
                "workList02": self.workList02 ,
                "workList03": self.workList03
            }

            fsaveNm = QFileDialog.getSaveFileName(self, "파일 저장")

            with open(fsaveNm[0], "w", encoding="utf-8") as file:
                json.dump(config_data, file, indent=4, ensure_ascii=False)



            print(config_data)
        except Exception as e:
            logger.info(f"saveConfig: {e}")

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
            logger.info("설정 파일이 없거나 형식이 잘못되었습니다.")

    # 설정값 로드 후 변수 재할당
    def refeshUiText(self):
        try:
            self.ui.plainTextEdit01.setPlainText(self.file01 )
            self.ui.plainTextEdit02.setPlainText(self.file02 )
            self.ui.plainTextEdit03.setPlainText(self.folder01 )
            self.ui.plainTextEdit04.setPlainText(self.saveNm01 )
            self.ui.plainTextEdit05.setPlainText(self.saveNm02 )

            self.table01_setData()
            self.table02_setData()
            self.table03_setData()
        except Exception as e:
            logger.info(f"refeshUiText: {e}")



    def setInitUiText(self):
        try:
            self.ui.plainTextEdit01.setPlainText(QCoreApplication.translate("MainWindow",msg01, None))
            self.ui.plainTextEdit02.setPlainText(QCoreApplication.translate("MainWindow",msg02, None))
            self.ui.plainTextEdit03.setPlainText(QCoreApplication.translate("MainWindow",msg03, None))
        except Exception as e:
            logger.info(f"setInitUiText: {e}")

    def getTextValues(self):
        try:
            self.file01 = self.ui.plainTextEdit01.toPlainText()
            self.file02 = self.ui.plainTextEdit02.toPlainText()
            self.folder01 = self.ui.plainTextEdit03.toPlainText()



            self.saveNm01 = self.ui.plainTextEdit04.toPlainText()
            self.saveNm02 = self.ui.plainTextEdit05.toPlainText()
            self.table01_saveData()
            self.table02_saveData()
            self.table03_saveData()
        except Exception as e:
            logger.info(f"getTextValues: {e}")



    def event_bttnExit(self):
        try:
            print("종료")
            QCoreApplication.instance().quit()
        except Exception as e:
            logger.info(f"event_bttnExit: {e}")

    def event_bttnLoad01(self):
        try:
            f =QFileDialog.getOpenFileName(self,'','','Excel(*.xlsx *xls)')
            if f[0]:
                self.ui.plainTextEdit01.setPlainText(f[0])
                self.file01 = f[0]
            else:
                self.ui.plainTextEdit01.setPlainText(msg01)
                print(msg01)
        except Exception as e:
            logger.info(f"event_bttnLoad01: {e}")


    def event_bttnLoad02(self):
        try:
            f =QFileDialog.getOpenFileName(self,'','','Excel(*.xlsx *xls)')
            if f[0]:
                self.ui.plainTextEdit02.setPlainText(f[0])
                self.file02 = f[0]
                print(f)
            else:
                self.ui.plainTextEdit02.setPlainText(msg02)
                print(msg02)
        except Exception as e:
            logger.info(f"event_bttnLoad02: {e}")


    def event_bttnLoad03(self):
        try:
            fd =QFileDialog.getExistingDirectory(self,'폴더선택','')
            if fd:
                self.ui.plainTextEdit03.setPlainText(fd)
                self.folder01 = fd
                print(fd)
            else:
                self.ui.plainTextEdit03.setPlainText(msg03)
                print(msg03)
        except Exception as e:
            logger.info(f"event_bttnLoad03: {e}")

    def event_table01_InsertRow(self):
        try:
            row_count = self.ui.table01.rowCount()
            self.ui.table01.insertRow(row_count)
        except Exception as e:
            logger.info(f"event_table01_InsertRow: {e}")

    def event_table02_InsertRow(self):
        try:
            row_count = self.ui.table02.rowCount()
            self.ui.table02.insertRow(row_count)
        except Exception as e:
            logger.info(f"event_table02_InsertRow: {e}")


    def table01_saveData(self):
        try:
            """테이블 데이터를 저장"""
            data = []
            row_count = self.ui.table01.rowCount()
            col_count = self.ui.table01.columnCount()

            for row in range(row_count):
                row_data = []
                for col in range(col_count):
                    item = self.ui.table01.item(row, col)
                    row_data.append(item.text().upper() if item else "")  # 빈 셀은 빈 문자열로 처리
                data.append(row_data)
            self.workList01 = data.copy()

            print("저장된 데이터:", data)  # 출력 (실제로는 파일 저장이나 DB 저장 가능)
        except Exception as e:
            logger.info(f"table01_saveData: {e}")


    def table02_saveData(self):
        try:
            """테이블 데이터를 저장"""
            data = []
            row_count = self.ui.table02.rowCount()
            col_count = self.ui.table02.columnCount()

            for row in range(row_count):
                for col in range(col_count):
                    item = self.ui.table02.item(row, col)
                    data.append(item.text())

            self.workList02 = data.copy()
        except Exception as e:
            logger.info(f"table02_saveData: {e}")


    def table03_saveData(self):
        try:
            """테이블 데이터를 저장"""
            data = []
            row_count = self.ui.table03.rowCount()
            col_count = self.ui.table03.columnCount()

            for row in range(row_count):
                for col in range(col_count):
                    item = self.ui.table03.item(row, col)
                    data.append(item.text().upper())

            self.workList03 = data.copy()
        except Exception as e:
            logger.info(f"table03_saveData: {e}")


    def table01_setData(self):
        try:
            """테이블에 초기 데이터를 설정하는 함수"""
            initial_data = self.workList01

            self.ui.table01.setRowCount(len(initial_data))  # 행 개수 설정

            for row, rowData in enumerate(initial_data):
                for col, value in enumerate(rowData):
                    self.ui.table01.setItem(row, col, QTableWidgetItem(value))
        except Exception as e:
            logger.info(f"table01_setData: {e}")

    def table02_setData(self):
        try:
            """테이블에 초기 데이터를 설정하는 함수"""
            initial_data = self.workList02

            self.ui.table02.setRowCount(len(initial_data))  # 행 개수 설정

            for row, rowData in enumerate(initial_data):
                self.ui.table02.setItem(row,0, QTableWidgetItem(rowData))
        except Exception as e:
            logger.info(f"table02_setData: {e}")



    def table03_setData(self):
        try:
            """테이블에 초기 데이터를 설정하는 함수"""
            initial_data = self.workList03

            self.ui.table03.setRowCount(len(initial_data))  # 행 개수 설정

            for row, rowData in enumerate(initial_data):
                self.ui.table03.setItem(row,0, QTableWidgetItem(rowData))

        except Exception as e:
            logger.info(f"table03_setData: {e}")





    def event_bttnStartWork(self):
        try:
            self.threadFlag = 1
            logger.info("작업을 시작합니다")
            self.getTextValues()
            self.worker_thread = WorkerThread()
            self.worker_thread.setValues( self.file01 ,self.file02 ,self.folder01 ,self.saveNm01 ,self.saveNm02,self.workList01,self.workList02,self.workList03 ,"N")


            self.worker_thread.start()
            self.event_showLogUi()  # 로그 UI 표시
        except Exception as e:
            logger.info(f"event_bttnStartWork: {e}")

    def event_bttnStartTest(self):
        try:
            self.threadFlag  = 1
            logger.info("작업 테스트를 시작합니다")
            self.getTextValues()
            self.worker_thread = WorkerThread()
            self.worker_thread.setValues(self.file01, self.file02, self.folder01, self.saveNm01, self.saveNm02,self.workList01, self.workList02, self.workList03 ,"Y")


            self.worker_thread.start()
            self.event_showLogUi()  # 로그 UI 표시
        except Exception as e:
            logger.info(f"event_bttnStartTest: {e}")


    def event_bttnForceStop(self):
        try:
            if self.threadFlag == 1:
                logger.info("작업을 강제로 종료합니다")

                writeFlag("0")
                logger.info(f"작업이 강제 종료되었습니다 {self.threadFlag}")
                self.threadFlag = 0
            else :
                logger.info("실행중인 작업이 없습니다")

                for proc in psutil.process_iter(['name']):
                    if proc.info['name'] and proc.info['name'].lower() == 'excel.exe':
                        proc.kill()
                        logger.info(f"엑셀 프로세스 PID {proc.pid} 종료 완료")



        except Exception as e:
            logger.info(f"event_bttnForceStop(): {e}")

    def event_showLogUi(self):
            self.log_window.show()

    def event_showUsageUi(self):
            self.usage_window.show()


if __name__ == "__main__":
    app = QApplication()
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
