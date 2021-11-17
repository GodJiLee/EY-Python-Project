import sys
import re
import datetime
from dateutil.relativedelta import relativedelta
import gc
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import pyodbc
import pandas as pd

class Calendar(QDialog):
    def __init__(self, parent):
        super(Calendar, self).__init__(parent)
        self.MyApp = MyApp

        self.setGeometry(500, 500, 400, 200)
        self.setWindowTitle("PyQt5 QCalendar")
        self.setWindowIcon(QIcon("python.png"))
        self.setWindowModality(Qt.NonModal)

        vbox = QVBoxLayout()
        self.calendar = QCalendarWidget()
        self.calendar.setGridVisible(True)

        self.label = QLabel("")
        self.label.setFont(QFont("Sanserif", 15))
        self.label.setStyleSheet('color:red')

        vbox.addWidget(self.calendar)
        vbox.addWidget(self.label)

        self.setLayout(vbox)

class Form(QWidget):
    def __init__(self):
        QWidget.__init__(self, flags=Qt.widget)
        self.setWindowTitle('QTreeWidget')
        self.setFixedWidth(210)
        self.setFixedHeight(150)

        ### QTreeView 생성 및 설정
        self.tw = QTreeWidget(self)
        self.tw.setColumnCount(2)
        self.tw.setHeaderLabels(['Account Type', 'Account Class'])
        self.tw.setAlternatingRowColors(True)
        self.tw.header().setSectionResizeMode(QHeaderView.Stretch)
        self.root = self.tw.invisibleRootItem()

        ## 데이터 계층적으로 저장하기
        data = [
            {"type": "1_Assets",
             "objects": [("11_유동자산"), ("12_비유동자산")]},
            {"type": "2_Liability",
             "objects": [("21_유동부채"), ("22_비유동부채")]}
        ]

        for d in data:
            parent = self.add_tree_root(d['type'], "")
            for child in d['objects']:
                self.add_tree_child(parent, *child)

        def add_tree_root(self, name: str, description: str):
            item = QTreeWidgetItem(self.tw)
            item.setText(0, name)
            item.setText(1, description)
            return item

        def add_tree_child(self, parent: QTreeWidgetItem, name: str, description: str):
            item = QTreeWidgetItem()
            item.setText(0, name)
            item.setText(1, description)
            parent.addChild(item)
            return item

        item = QTreeWidgetItem()
        item.setText(0, "1_Assets")
        sub_item = QTreeWidgetItem()

        sub_item.setText(0, "11_유동자산")
        sub_item.setText(1, "12_비유동자산")

        item.setText(1, "2_Liability")
        sub_item.setText(0, "21_유동부채")
        sub_item.setText(1, "22_비유동부채")

        item.addChild(sub_item)
        self.root.addChild(item)

        self.root.addChild(item)

class DataFrameModel(QAbstractTableModel):
    DtypeRole = Qt.UserRole + 1000
    ValueRole = Qt.UserRole + 1001

    def __init__(self, df=pd.DataFrame(), parent=None):
        super(DataFrameModel, self).__init__(parent)
        self._dataframe = df

    def setDataFrame(self, dataframe):
        self.beginResetModel()
        self._dataframe = dataframe.copy()
        self.endResetModel()

    def dataFrame(self):
        return self._dataframe

    dataFrame = pyqtProperty(pd.DataFrame, fget=dataFrame, fset=setDataFrame)

    @pyqtSlot(int, Qt.Orientation, result=str)
    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return self._dataframe.columns[section]
            else:
                return str(self._dataframe.index[section])
        return QVariant()

    def rowCount(self, parent=QModelIndex()):
        if parent.isValid():
            return 0
        return len(self._dataframe.index)

    def columnCount(self, parent=QModelIndex()):
        if parent.isValid():
            return 0
        return self._dataframe.columns.size

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid() or not (0 <= index.row() < self.rowCount() and 0 <= index.column() < self.columnCount()):
            return QVariant()
        row = self._dataframe.index[index.row()]
        col = self._dataframe.columns[index.column()]
        dt = self._dataframe[col].dtype

        val = self._dataframe.iloc[row][col]
        if role == Qt.DisplayRole:
            return str(val)
        elif role == DataFrameModel.ValueRole:
            return val
        if role == DataFrameModel.DtypeRole:
            return dt
        return QVariant()

    def roleNames(self):
        roles = {
            Qt.DisplayRole: b'display',
            DataFrameModel.DtypeRole: b'dtype',
            DataFrameModel.ValueRole: b'value'
        }
        return roles


class ListBoxWidget(QListWidget):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.resize(600, 600)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls:
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()

            links = []
            for url in event.mimeData().urls():

                if url.isLocalFile():
                    links.append(str(url.toLocalFile()))
                else:
                    links.append(str(url.toString()))
            self.addItems(links)
        else:
            event.ignore()




class MyApp(QWidget):

    def __init__(self):
        super().__init__()
        self.init_UI()
        ##Initialize Variables
        self.selected_project_id = None
        self.selected_server_name = "--서버 목록--"
        self.dataframe = None
        self.cnxn = None
        self.selected_scenario_class_index = 0
        self.selected_scenario_subclass_index = 0
        self.scenario_dic = {}
        self.selected_scenario_group = None
        self.SaveRoute = None


    def MessageBox_Open(self, text):
        self.msg = QMessageBox()
        self.msg.setIcon(QMessageBox.Information)
        self.msg.setWindowTitle("Warning")
        self.msg.setWindowIcon(QIcon("./EY_logo.png"))
        self.msg.setText(text)
        self.msg.exec_()


    def alertbox_open(self):
        self.alt = QMessageBox()
        self.alt.setIcon(QMessageBox.Information)
        self.alt.setWindowTitle('필수 입력값 누락')
        self.alt.setWindowIcon(QIcon("./EY_logo.png"))
        self.alt.setText('필수 입력값이 누락되었습니다.')
        self.alt.exec_()


    def alertbox_open2(self, state):
        self.alt = QMessageBox()
        self.alt.setIcon(QMessageBox.Information)
        txt = state
        self.alt.setWindowTitle('필수 입력값 타입 오류')
        self.alt.setWindowIcon(QIcon("./EY_logo.png"))
        self.alt.setText(txt + ' 값을 ' + '숫자로만 입력해주시기 바랍니다.')
        self.alt.exec_()


    def init_UI(self):

        image = QImage('./dark_gray.png')
        scaledImg = image.scaled(QSize(1000, 900))
        palette = QPalette()
        palette.setBrush(10, QBrush(scaledImg))
        self.setPalette(palette)

        pixmap = QPixmap('./title.png')
        lbl_img = QLabel()
        lbl_img.setPixmap(pixmap)
        lbl_img.setStyleSheet("background-color:#000000")

        widget_layout = QBoxLayout(QBoxLayout.TopToBottom)
        self.splitter_layout = QSplitter(Qt.Vertical)

        self.splitter_layout.addWidget(lbl_img)
        self.splitter_layout.addWidget(self.Connect_ServerInfo_Group())
        self.splitter_layout.addWidget(self.Show_DataFrame_Group())
        self.splitter_layout.addWidget(self.Save_Buttons_Group())
        self.splitter_layout.setHandleWidth(0)

        widget_layout.addWidget(self.splitter_layout)
        self.setLayout(widget_layout)

        self.setWindowIcon(QIcon("./EY_logo.png"))
        self.setWindowTitle('Scenario')

        self.setGeometry(300, 100, 1000, 900)
        self.show()



    def connectButtonClicked(self):

        password = ''
        ecode = self.line_ecode.text().strip() ##leading/trailing space 포함시 제거
        user = 'guest'
        server = self.selected_server_name
        db = 'master'

        # 예외처리 - 서버 선택
        if server == "--서버 목록--":
            self.MessageBox_Open("서버가 선택되어 있지 않습니다")
            return

        # 예외처리 - Ecode 이상
        elif ecode.isdigit() is False:
            self.MessageBox_Open("Engagement Code가 잘못되었습니다")
            self.ProjectCombobox.clear()
            self.ProjectCombobox.addItem("프로젝트가 없습니다")
            return


        server_path = f"DRIVER={{SQL Server}};SERVER={server};uid={user};pwd={password};DATABASE={db};trusted_connection=yes"

        # 예외처리 - 접속 정보 오류
        try:
            self.cnxn = pyodbc.connect(server_path)
        except:
            QMessageBox.about(self, "Warning", "접속 정보가 잘못되었습니다")
            return

        cursor = self.cnxn.cursor()

        sql_query = f"""
                           SELECT ProjectName
                           From [DataAnalyticsRepository].[dbo].[Projects]
                           WHERE EngagementCode IN ({ecode})
                           AND DeletedBy IS NULL
                     """


        try:
            selected_project_names = pd.read_sql(sql_query, self.cnxn)

        except:
            self.MessageBox_Open("Engagement Code를 입력하세요.")
            self.ProjectCombobox.clear()
            self.ProjectCombobox.addItem("프로젝트가 없습니다")
            return

        # 예외처리 - 해당 ecode에 아무런 프로젝트가 없는 경우
        if len(selected_project_names) == 0:
            self.MessageBox_Open("해당 Engagement Code 내 프로젝트가 존재하지 않습니다.")
            self.ProjectCombobox.clear()
            self.ProjectCombobox.addItem("프로젝트가 없습니다.")
            return


        ## 서버 연결 시 - 기존 저장 정보를 초기화(메모리 관리)
        del self.selected_project_id, self.dataframe, self.scenario_dic
        gc.collect()

        self.ProjectCombobox.clear()
        self.ProjectCombobox.addItem("--프로젝트 목록--")
        for i in range(len(selected_project_names)):
            self.ProjectCombobox.addItem(selected_project_names.iloc[i, 0])

        self.selected_project_id = None
        self.dataframe = None
        self.scenario_dic = {}

    def Server_ComboBox_Selected(self, text):
        self.selected_server_name = text

    def Project_ComboBox_Selected(self, text):
        ecode = self.line_ecode.text().strip()  # leading/trailing space 제거
        pname = text

        ## 예외처리 - 서버가 연결되지 않은 상태로 Project name Combo box를 건드리는 경우
        if self.cnxn is None:
            return

        cursor = self.cnxn.cursor()

        sql_query = f"""
                        SELECT Project_ID
                        FROM [DataAnalyticsRepository].[dbo].[Projects]
                        WHERE ProjectName IN (\'{pname}\')
                        AND EngagementCode IN ({ecode})
                        AND DeletedBy is Null
                     """

        ## 예외처리 - 에러 표시인 "프로젝트가 없습니다"가 선택되어 있는 경우
        try:
            self.selected_project_id = pd.read_sql(sql_query, self.cnxn).iloc[0, 0]
        except:
            self.selected_project_id = None



    def Connect_ServerInfo_Group(self):

        groupbox = QGroupBox('접속 정보')
        self.setStyleSheet('QGroupBox  {color: white;}')
        font_groupbox = groupbox.font()
        font_groupbox.setBold(True)
        groupbox.setFont(font_groupbox)

        ##labels 생성 및 스타일 지정
        label1 = QLabel('Server : ', self)
        label2 = QLabel('Engagement Code : ', self)
        label3 = QLabel('Project Name : ', self)
        label4 = QLabel('Scenario : ', self)

        font1 = label1.font()
        font1.setBold(True)
        label1.setFont(font1)
        font2 = label2.font()
        font2.setBold(True)
        label2.setFont(font2)
        font3 = label3.font()
        font3.setBold(True)
        label3.setFont(font3)
        font4 = label4.font()
        font4.setBold(True)
        label4.setFont(font4)

        label1.setStyleSheet("color: white;")
        label2.setStyleSheet("color: white;")
        label3.setStyleSheet("color: white;")
        label4.setStyleSheet("color: white;")


        ##서버 선택 콤보박스
        self.cb_server = QComboBox(self)
        self.cb_server.addItem('--서버 목록--')
        for i in range(1, 9):
            self.cb_server.addItem(f'KRSEOVMPPACSQ0{i}\INST1')


        ##scenario 유형 콤보박스
        self.comboBig = QComboBox(self)

        self.comboBig.addItem('Data 완전성', ['--시나리오 목록--', '04 : 계정 사용빈도 N번이하인 계정이 포함된 전표리스트', '05 : 당기 생성된 계정리스트 추출'])
        self.comboBig.addItem('Data Timing',
                              ['--시나리오 목록--', '06 : 결산일 전후 T일 입력 전표', '07 : 영업일 전기/입력 전표', '08 : 효력, 입력 일자 간 차이가 N일 이상인 전표'])
        self.comboBig.addItem('Data 업무분장',
                              ['--시나리오 목록--', '09 : 전표 작성 빈도수가 N회 이하인 작성자에 의한 생성된 전표', '10 : 특정 전표 입력자(W)에 의해 생성된 전표'])
        self.comboBig.addItem('Data 분개검토',
                              ['--시나리오 목록--', '11 : 특정한 주계정(A)과 특정한 상대계정(B)이 아닌 전표리스트 검토', '12 : 특정 계정(A)이 감소할 때 상대계정 리스트 검토'])
        self.comboBig.addItem('기타', ['--시나리오 목록--', '13 : 연속된 숫자로 끝나는 금액 검토',
                                     '14 : 전표 description에 공란 또는 특정단어(key word)가 입력되어 있는 전표 리스트 (TE금액 제시 가능)'])

        ##시나리오 세부 내역/프로젝트 선택 콤보박스
        self.comboSmall = QComboBox(self)
        self.comboSmall.addItems(self.comboBig.itemData(0))

        self.ProjectCombobox = QComboBox(self)

        ##Engagement code 입력 line
        self.line_ecode = QLineEdit(self)
        self.line_ecode.setText("")

        ##SQL SERVER CONNECT 버튼 생성 및 스타일 지정
        btn_server = QPushButton('   SQL Server Connect', self)
        font_btn_server = btn_server.font()
        font_btn_server.setBold(True)
        btn_server.setFont(font_btn_server)
        btn_server.setStyleSheet('color:white;  background-image : url(./bar.png)')

        ##Input Conditions 버튼 생성 및 스타일 지정
        btn_condition = QPushButton('   Input Conditions', self)
        font_btn_condition = btn_condition.font()
        font_btn_condition.setBold(True)
        btn_condition.setFont(font_btn_condition)
        btn_condition.setStyleSheet('color:white;  background-image : url(./bar.png)')

        ##Signal 함수들
        self.comboBig.activated[str].connect(self.ComboBig_Selected)
        self.comboSmall.activated[str].connect(self.ComboSmall_Selected)
        self.cb_server.activated[str].connect(self.Server_ComboBox_Selected)
        btn_server.clicked.connect(self.connectButtonClicked)
        self.ProjectCombobox.activated[str].connect(self.Project_ComboBox_Selected)
        btn_condition.clicked.connect(self.connectDialog)

        ##layout 쌓기
        grid = QGridLayout()
        grid.addWidget(label1, 0, 0)
        grid.addWidget(label2, 1, 0)
        grid.addWidget(label3, 2, 0)
        grid.addWidget(label4, 3, 0)
        grid.addWidget(self.cb_server, 0, 1)
        grid.addWidget(btn_server, 0, 2)
        grid.addWidget(self.comboBig, 3, 1)
        grid.addWidget(self.comboSmall, 4, 1)
        grid.addWidget(btn_condition, 4, 2)
        grid.addWidget(self.line_ecode, 1, 1)
        grid.addWidget(self.ProjectCombobox, 2, 1)

        groupbox.setLayout(grid)
        return groupbox


    def ComboBig_Selected(self, text):
        idx = self.comboBig.currentIndex()
        self.selected_scenario_class_index = idx
        self.comboSmall.clear()
        self.comboSmall.addItems(self.comboBig.itemData(idx))


    def ComboSmall_Selected(self, text):
        self.selected_scenario_subclass_index = self.comboSmall.currentIndex()


    def connectDialog(self):
        if self.cnxn is None:
            self.MessageBox_Open("SQL 서버가 연결되어 있지 않습니다")
            return

        elif self.selected_project_id is None:
            self.MessageBox_Open("프로젝트가 선택되지 않았습니다")
            return

        elif self.selected_scenario_subclass_index == 0:
            self.MessageBox_Open("시나리오가 선택되지 않았습니다")
            return

        if self.selected_scenario_class_index == 0 and self.selected_scenario_subclass_index == 1:
            self.Dialog4()

        elif self.selected_scenario_class_index == 0 and self.selected_scenario_subclass_index == 2:
            self.Dialog5()

        elif self.selected_scenario_class_index == 1 and self.selected_scenario_subclass_index == 1:
            self.Dialog6()

        elif self.selected_scenario_class_index == 1 and self.selected_scenario_subclass_index == 2:
            self.Dialog7()

        elif self.selected_scenario_class_index == 1 and self.selected_scenario_subclass_index == 3:
            self.Dialog8()

        elif self.selected_scenario_class_index == 2 and self.selected_scenario_subclass_index == 1:
            self.Dialog9()

        elif self.selected_scenario_class_index == 2 and self.selected_scenario_subclass_index == 2:
            self.Dialog10()

        elif self.selected_scenario_class_index == 3 and self.selected_scenario_subclass_index == 1:
            self.Dialog11()

        elif self.selected_scenario_class_index == 3 and self.selected_scenario_subclass_index == 2:
            self.Dialog12()

        elif self.selected_scenario_class_index == 4 and self.selected_scenario_subclass_index == 1:
            self.Dialog13()

        elif self.selected_scenario_class_index == 4 and self.selected_scenario_subclass_index == 2:
            self.Dialog14()




    def Dialog4(self):
        self.dialog4 = QDialog()
        self.dialog4.setStyleSheet('background-color: #2E2E38')
        self.dialog4.setWindowIcon(QIcon('./EY_logo.png'))

        self.btn2 = QPushButton('   Extract Data', self.dialog4)
        self.btn2.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked4)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        self.btnDialog = QPushButton('   Close', self.dialog4)
        self.btnDialog.setStyleSheet(
            'color:white;  background-image : url(./bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close4)

        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)

        self.btn2.resize(110, 30)
        self.btnDialog.resize(110, 30)

        label_freq = QLabel('사용빈도(N)* :', self.dialog4)
        label_freq.setStyleSheet('color: white;')

        font1 = label_freq.font()
        font1.setBold(True)
        label_freq.setFont(font1)

        self.D4_N = QLineEdit(self.dialog4)
        self.D4_N.setStyleSheet('background-color: white;')

        label_TE = QLabel('중요성금액: ', self.dialog4)
        label_TE.setStyleSheet('color: white;')

        font2 = label_TE.font()
        font2.setBold(True)
        label_TE.setFont(font2)

        self.D4_TE = QLineEdit(self.dialog4)
        self.D4_TE.setStyleSheet('background-color: white;')

        self.D4_N.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D4_TE.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        layout1 = QGridLayout()
        layout1.addWidget(label_freq, 0, 0)
        layout1.addWidget(self.D4_N, 0, 1)
        layout1.addWidget(label_TE, 1, 0)
        layout1.addWidget(self.D4_TE, 1, 1)

        layout2 = QHBoxLayout()
        layout2.addStretch()
        layout2.addStretch()
        layout2.addWidget(self.btn2)
        layout2.addWidget(self.btnDialog)

        layout2.setContentsMargins(-1, 10, -1, -1)

        main_layout = QVBoxLayout()
        main_layout.setAlignment(Qt.AlignTop)
        main_layout.addLayout(layout1)
        main_layout.addLayout(layout2)

        self.dialog4.setLayout(main_layout)
        self.dialog4.setGeometry(300, 300, 500, 150)

        self.dialog4.setWindowTitle('Scenario4')
        self.dialog4.setWindowModality(Qt.NonModal)
        self.dialog4.show()

    def Dialog5(self):  # QDialog 창 수정하지 말 것
        self.dialog5 = QDialog()
        self.dialog5.setStyleSheet('background-color: #2E2E38')
        self.dialog5.setWindowIcon(QIcon('./EY_logo.png'))

        ### Extract Data
        self.btn2 = QPushButton(' Extract Data', self.dialog5)
        self.btn2.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked5_Non_SAP)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        self.btn3 = QPushButton(' Extract Data', self.dialog5)
        self.btn3.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btn3.clicked.connect(self.extButtonClicked5_SAP)

        font11 = self.btn3.font()
        font11.setBold(True)
        self.btn3.setFont(font11)

        ### Close
        self.btnDialog = QPushButton('Close', self.dialog5)
        self.btnDialog.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close5)

        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)

        self.btnDialog1 = QPushButton('Close', self.dialog5)
        self.btnDialog1.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btnDialog1.clicked.connect(self.dialog_close5)

        font13 = self.btnDialog1.font()
        font13.setBold(True)
        self.btnDialog1.setFont(font13)

        ### 라벨1 - 계정코드 입력
        label_AccCode = QLabel('Enter your Account Code: ', self.dialog5)
        label_AccCode.setStyleSheet('color: white;')
        label_AccCode.setFont(QFont('Arial', 12))

        font1 = label_AccCode.font()
        font1.setBold(True)
        label_AccCode.setFont(font1)

        ### 라벨 2 - 입력 예시
        label_Example = QLabel('※ 입력 예시: OO', self.dialog5)
        label_Example.setStyleSheet('color: red;')
        label_Example.setFont(QFont('Times font', 9))

        font2 = label_Example.font()
        font2.setBold(False)
        label_Example.setFont(font2)

        label_SAP_Example = QLabel('※ SKA1 파일을 Drop 하십시오', self.dialog5)
        label_SAP_Example.setStyleSheet('color: red;')
        label_SAP_Example.setFont(QFont('Times font', 9))

        font12 = label_SAP_Example.font()
        font12.setBold(False)
        label_SAP_Example.setFont(font12)

        ### TextEdit - 계정코드 Paste
        self.MyInput = QTextEdit(self.dialog5)
        self.MyInput.setAcceptRichText(False)
        self.MyInput.setStyleSheet('background-color: white;')

        ### ListBox Widget
        self.listbox_drops = ListBoxWidget()
        self.listbox_drops.setStyleSheet('background-color: white;')

        layout = QVBoxLayout()

        layout1 = QVBoxLayout()
        sublayout1 = QVBoxLayout()
        sublayout2 = QHBoxLayout()

        layout2 = QVBoxLayout()
        sublayout3 = QVBoxLayout()
        sublayout4 = QHBoxLayout()

        tab1 = QWidget()
        tab2 = QWidget()
        tabs = QTabWidget()

        sublayout1.addWidget(label_AccCode)
        sublayout1.addWidget(label_Example)
        sublayout1.addWidget(self.MyInput)

        sublayout2.addStretch(1)
        sublayout2.addWidget(self.btn2, stretch=1, alignment=Qt.AlignBottom)
        sublayout2.addWidget(self.btnDialog, stretch=1, alignment=Qt.AlignBottom)
        sublayout2.addStretch(1)

        layout1.addLayout(sublayout1, stretch=4)
        layout1.addLayout(sublayout2, stretch=1)

        sublayout3.addWidget(label_SAP_Example)
        sublayout3.addWidget(self.listbox_drops)

        sublayout4.addStretch(1)
        sublayout4.addWidget(self.btn3, stretch=1, alignment=Qt.AlignBottom)
        sublayout4.addWidget(self.btnDialog1, stretch=1, alignment=Qt.AlignBottom)
        sublayout4.addStretch(1)

        layout2.addLayout(sublayout3, stretch=4)
        layout2.addLayout(sublayout4, stretch=1)

        tab1.setLayout(layout1)
        tab2.setLayout(layout2)

        tabs.addTab(tab1, "Non-SAP")
        tabs.addTab(tab2, "SAP")

        layout.addWidget(tabs)

        self.dialog5.setLayout(layout)

        self.dialog5.resize(465, 400)

        self.dialog5.setWindowTitle('Scenario5')
        self.dialog5.setWindowModality(Qt.NonModal)
        self.dialog5.show()

    def Dialog6(self):
        self.dialog6 = QDialog()
        self.dialog6.setStyleSheet('background-color: #2E2E38')
        self.dialog6.setWindowIcon(QIcon("./EY_logo.png"))

        self.btn2 = QPushButton('   Extract Data', self.dialog6)
        self.btn2.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked6)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        self.btnDialog = QPushButton("   Close", self.dialog6)
        self.btnDialog.setStyleSheet(
            'color:white;  background-image : url(./bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close6)

        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)

        self.btn2.resize(110, 30)
        self.btnDialog.resize(110, 30)

        labelDate = QLabel('결산일* : ', self.dialog6)
        labelDate.setStyleSheet("color: white;")

        font1 = labelDate.font()
        font1.setBold(True)
        labelDate.setFont(font1)

        self.D6_Date = QLineEdit(self.dialog6)
        self.D6_Date.setStyleSheet("background-color: white;")
        self.D6_Date.setPlaceholderText('날짜를 선택하세요')

        self.btnDate = QPushButton("Date", self.dialog6)
        self.btnDate.resize(65, 22)
        self.new_calendar = Calendar(self)
        self.new_calendar.calendar.clicked.connect(self.handle_date_clicked)
        self.btnDate.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btnDate.clicked.connect(self.calendar)
        font11 = self.btnDate.font()
        font11.setBold(True)
        self.btnDate.setFont(font11)

        labelDate2 = QLabel('T일* : ', self.dialog6)
        labelDate2.setStyleSheet("color: white;")

        font2 = labelDate2.font()
        font2.setBold(True)
        labelDate2.setFont(font2)

        self.D6_Date2 = QLineEdit(self.dialog6)
        self.D6_Date2.setStyleSheet("background-color: white;")
        self.D6_Date2.setPlaceholderText('T 값을 입력하세요')

        label_tree = QLabel('특정 계정명 : ', self.dialog6)
        label_tree.setStyleSheet("color: white;")
        font4 = label_tree.font()
        font4.setBold(True)
        label_tree.setFont(font4)

        # 트리 예시
        self.account_tree = QTreeWidget(self.dialog6)
        self.account_tree.setColumnCount(2)
        self.account_tree.setStyleSheet("background-color: white;")
        self.account_tree.setHeaderLabels(['Account Type'])
        self.account_tree.setAlternatingRowColors(False)
        self.account_tree.header().setVisible(True)

        itemTop1 = QTreeWidgetItem(self.account_tree)
        itemTop1.setText(0, "1_Assets")
        itemTop1.setFlags(itemTop1.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)

        itemChild1 = QTreeWidgetItem(itemTop1)
        itemChild1.setFlags(itemChild1.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
        itemChild1.setText(0, '11_유동자산')
        itemChild1.setCheckState(0, Qt.Unchecked)

        itemChild11 = QTreeWidgetItem(itemChild1)
        itemChild11.setFlags(itemChild11.flags() | Qt.ItemIsUserCheckable)
        itemChild11.setText(0, '1101_현금및현금성자산')
        itemChild11.setCheckState(0, Qt.Unchecked)

        itemChild12 = QTreeWidgetItem(itemChild1)
        itemChild12.setFlags(itemChild12.flags() | Qt.ItemIsUserCheckable)
        itemChild12.setText(0, '1105_매출채권')
        itemChild12.setCheckState(0, Qt.Unchecked)

        itemChild2 = QTreeWidgetItem(itemTop1)
        itemChild2.setFlags(itemChild2.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
        itemChild2.setText(0, '12_비유동자산')
        itemChild2.setCheckState(0, Qt.Unchecked)

        itemTop2 = QTreeWidgetItem(self.account_tree)
        itemTop2.setText(0, '2_Liability')
        itemTop2.setFlags(itemTop2.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)

        itemChild3 = QTreeWidgetItem(itemTop2)
        itemChild3.setFlags(itemChild3.flags() | Qt.ItemIsUserCheckable)
        itemChild3.setText(0, '21_유동부채')
        itemChild3.setCheckState(0, Qt.Unchecked)

        itemChild4 = QTreeWidgetItem(itemTop2)
        itemChild4.setFlags(itemChild4.flags() | Qt.ItemIsUserCheckable)
        itemChild4.setText(0, '22_비유동부채')
        itemChild4.setCheckState(0, Qt.Unchecked)

        labelJE = QLabel('전표입력자 : ', self.dialog6)
        labelJE.setStyleSheet("color: white;")

        font4 = labelJE.font()
        font4.setBold(True)
        labelJE.setFont(font4)

        self.D6_JE = QLineEdit(self.dialog6)
        self.D6_JE.setStyleSheet("background-color: white;")
        self.D6_JE.setPlaceholderText('전표입력자 ID를 입력하세요')

        labelCost = QLabel('중요성금액 : ', self.dialog6)
        labelCost.setStyleSheet("color: white;")

        font5 = labelCost.font()
        font5.setBold(True)
        labelCost.setFont(font5)

        self.D6_Cost = QLineEdit(self.dialog6)
        self.D6_Cost.setStyleSheet("background-color: white;")
        self.D6_Cost.setPlaceholderText('100,000,000원 이상 입력하세요')

        self.D6_Date.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D6_Date2.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D6_JE.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D6_Cost.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        layout1 = QGridLayout()
        layout1.addWidget(labelDate, 0, 0)
        layout1.addWidget(self.D6_Date, 0, 1)
        layout1.addWidget(self.btnDate, 0, 2)
        layout1.addWidget(labelDate2, 1, 0)
        layout1.addWidget(self.D6_Date2, 1, 1)
        layout1.addWidget(label_tree, 2, 0)
        layout1.addWidget(self.account_tree, 2, 1)
        layout1.addWidget(labelJE, 3, 0)
        layout1.addWidget(self.D6_JE, 3, 1)
        layout1.addWidget(labelCost, 4, 0)
        layout1.addWidget(self.D6_Cost, 4, 1)

        layout2 = QHBoxLayout()
        layout2.addStretch()
        layout2.addStretch()
        layout2.addWidget(self.btn2)
        layout2.addWidget(self.btnDialog)

        layout2.setContentsMargins(-1, 10, -1, -1)

        main_layout = QVBoxLayout()
        main_layout.setAlignment(Qt.AlignTop)
        main_layout.addLayout(layout1)
        main_layout.addLayout(layout2)

        self.dialog6.setLayout(main_layout)
        self.dialog6.setGeometry(300, 300, 700, 400)

        self.dialog6.setWindowTitle("Scenario6")
        self.dialog6.setWindowModality(Qt.NonModal)
        self.dialog6.show()

    def Dialog7(self):
        self.dialog7 = QDialog()
        self.dialog7.setStyleSheet('background-color: #2E2E38')
        self.dialog7.setWindowIcon(QIcon("./EY_logo.png"))

        self.btn2 = QPushButton('   Extract Data', self.dialog7)
        self.btn2.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked7)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        self.btnDialog = QPushButton("   Close", self.dialog7)
        self.btnDialog.setStyleSheet(
            'color:white;  background-image : url(./bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close7)

        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)

        self.btn2.resize(110, 30)
        self.btnDialog.resize(110, 30)

        self.rbtn1 = QRadioButton('Effective Date', self.dialog7)
        self.rbtn1.setStyleSheet("color: white;")

        font1 = self.rbtn1.font()
        font1.setBold(True)
        self.rbtn1.setFont(font1)

        self.rbtn1.setChecked(True)

        self.rbtn2 = QRadioButton('Entry Date', self.dialog7)
        self.rbtn2.setStyleSheet("color: white;")

        font2 = self.rbtn2.font()
        font2.setBold(True)
        self.rbtn2.setFont(font2)

        labelDate = QLabel('Effective Date/Entry Date* : ', self.dialog7)
        labelDate.setStyleSheet("color: white;")

        font3 = labelDate.font()
        font3.setBold(True)
        labelDate.setFont(font3)

        self.D7_Date = QLineEdit(self.dialog7)
        self.D7_Date.setStyleSheet("background-color: white;")
        self.D7_Date.setPlaceholderText('날짜를 선택하세요')

        self.btnDate = QPushButton("Date", self.dialog7)
        self.btnDate.resize(65, 22)
        self.new_calendar = Calendar(self)
        self.new_calendar.calendar.clicked.connect(self.handle_date_clicked2)
        self.btnDate.setStyleSheet(
            'color:white;  background-image : url(./bar.png)')
        self.btnDate.clicked.connect(self.calendar)

        font11 = self.btnDate.font()
        font11.setBold(True)
        self.btnDate.setFont(font11)

        label_tree = QLabel('특정 계정명 : ', self.dialog7)
        label_tree.setStyleSheet("color: white;")
        font4 = label_tree.font()
        font4.setBold(True)
        label_tree.setFont(font4)

        # 트리 예시
        self.account_tree = QTreeWidget(self.dialog7)
        self.account_tree.setColumnCount(2)
        self.account_tree.setStyleSheet("background-color: white;")
        self.account_tree.setHeaderLabels(['Account Type'])
        self.account_tree.setAlternatingRowColors(False)
        self.account_tree.header().setVisible(True)

        itemTop1 = QTreeWidgetItem(self.account_tree)
        itemTop1.setText(0, "1_Assets")
        itemTop1.setFlags(itemTop1.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)

        itemChild1 = QTreeWidgetItem(itemTop1)
        itemChild1.setFlags(itemChild1.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
        itemChild1.setText(0, '11_유동자산')
        itemChild1.setCheckState(0, Qt.Unchecked)

        itemChild11 = QTreeWidgetItem(itemChild1)
        itemChild11.setFlags(itemChild11.flags() | Qt.ItemIsUserCheckable)
        itemChild11.setText(0, '1101_현금및현금성자산')
        itemChild11.setCheckState(0, Qt.Unchecked)

        itemChild12 = QTreeWidgetItem(itemChild1)
        itemChild12.setFlags(itemChild12.flags() | Qt.ItemIsUserCheckable)
        itemChild12.setText(0, '1105_매출채권')
        itemChild12.setCheckState(0, Qt.Unchecked)

        itemChild2 = QTreeWidgetItem(itemTop1)
        itemChild2.setFlags(itemChild2.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
        itemChild2.setText(0, '12_비유동자산')
        itemChild2.setCheckState(0, Qt.Unchecked)

        itemTop2 = QTreeWidgetItem(self.account_tree)
        itemTop2.setText(0, '2_Liability')
        itemTop2.setFlags(itemTop2.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)

        itemChild3 = QTreeWidgetItem(itemTop2)
        itemChild3.setFlags(itemChild3.flags() | Qt.ItemIsUserCheckable)
        itemChild3.setText(0, '21_유동부채')
        itemChild3.setCheckState(0, Qt.Unchecked)

        itemChild4 = QTreeWidgetItem(itemTop2)
        itemChild4.setFlags(itemChild4.flags() | Qt.ItemIsUserCheckable)
        itemChild4.setText(0, '22_비유동부채')
        itemChild4.setCheckState(0, Qt.Unchecked)

        labelJE = QLabel('전표입력자 : ', self.dialog7)
        labelJE.setStyleSheet("color: white;")

        font5 = labelJE.font()
        font5.setBold(True)
        labelJE.setFont(font5)

        self.D7_JE = QLineEdit(self.dialog7)
        self.D7_JE.setStyleSheet("background-color: white;")
        self.D7_JE.setPlaceholderText('전표입력자 ID를 입력하세요')

        labelCost = QLabel('중요성금액 : ', self.dialog7)
        labelCost.setStyleSheet("color: white;")

        font6 = labelCost.font()
        font6.setBold(True)
        labelCost.setFont(font6)

        self.D7_Cost = QLineEdit(self.dialog7)
        self.D7_Cost.setStyleSheet("background-color: white;")
        self.D7_Cost.setPlaceholderText('100,000,000원 이상 입력하세요')

        self.D7_Date.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D7_JE.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D7_Cost.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        layout0 = QGridLayout()
        layout0.addWidget(self.rbtn1, 0, 0)
        layout0.addWidget(self.rbtn2, 0, 1)

        layout1 = QGridLayout()
        layout1.addWidget(labelDate, 0, 0)
        layout1.addWidget(self.D7_Date, 0, 1)
        layout1.addWidget(self.btnDate, 0, 2)
        layout1.addWidget(label_tree, 1, 0)
        layout1.addWidget(self.account_tree, 1, 1)
        layout1.addWidget(labelJE, 2, 0)
        layout1.addWidget(self.D7_JE, 2, 1)
        layout1.addWidget(labelCost, 3, 0)
        layout1.addWidget(self.D7_Cost, 3, 1)

        layout2 = QHBoxLayout()
        layout2.addStretch()
        layout2.addStretch()
        layout2.addWidget(self.btn2)
        layout2.addWidget(self.btnDialog)

        layout2.setContentsMargins(-1, 10, -1, -1)

        main_layout = QVBoxLayout()
        main_layout.setAlignment(Qt.AlignTop)
        main_layout.addLayout(layout0)
        main_layout.addLayout(layout1)
        main_layout.addLayout(layout2)

        self.dialog7.setLayout(main_layout)
        self.dialog7.setGeometry(300, 300, 700, 400)
        self.dialog7.setWindowTitle("Scenario7")
        self.dialog7.setWindowModality(Qt.NonModal)
        self.dialog7.show()

    def Dialog8(self):
        self.dialog8 = QDialog()
        self.dialog8.setStyleSheet('background-color: #2E2E38')
        self.dialog8.setWindowIcon(QIcon("./EY_logo.png"))

        self.btn2 = QPushButton('   Extract Data', self.dialog8)
        self.btn2.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked8)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        self.btnDialog = QPushButton("   Close", self.dialog8)
        self.btnDialog.setStyleSheet(
            'color:white;  background-image : url(./bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close8)

        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)

        self.btn2.resize(110, 30)
        self.btnDialog.resize(110, 30)

        labelDate = QLabel('N일* : ', self.dialog8)
        labelDate.setStyleSheet("color: white;")

        font1 = labelDate.font()
        font1.setBold(True)
        labelDate.setFont(font1)

        self.D8_N = QLineEdit(self.dialog8)
        self.D8_N.setStyleSheet("background-color: white;")
        self.D8_N.setPlaceholderText('N 값을 입력하세요')

        label_tree = QLabel('특정 계정명 : ', self.dialog8)
        label_tree.setStyleSheet("color: white;")
        font4 = label_tree.font()
        font4.setBold(True)
        label_tree.setFont(font4)

        # 트리 예시
        self.account_tree = QTreeWidget(self.dialog8)
        self.account_tree.setColumnCount(2)
        self.account_tree.setStyleSheet("background-color: white;")
        self.account_tree.setHeaderLabels(['Account Type'])
        self.account_tree.setAlternatingRowColors(False)
        self.account_tree.header().setVisible(True)

        itemTop1 = QTreeWidgetItem(self.account_tree)
        itemTop1.setText(0, "1_Assets")
        itemTop1.setFlags(itemTop1.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)

        itemChild1 = QTreeWidgetItem(itemTop1)
        itemChild1.setFlags(itemChild1.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
        itemChild1.setText(0, '11_유동자산')
        itemChild1.setCheckState(0, Qt.Unchecked)

        itemChild11 = QTreeWidgetItem(itemChild1)
        itemChild11.setFlags(itemChild11.flags() | Qt.ItemIsUserCheckable)
        itemChild11.setText(0, '1101_현금및현금성자산')
        itemChild11.setCheckState(0, Qt.Unchecked)

        itemChild12 = QTreeWidgetItem(itemChild1)
        itemChild12.setFlags(itemChild12.flags() | Qt.ItemIsUserCheckable)
        itemChild12.setText(0, '1105_매출채권')
        itemChild12.setCheckState(0, Qt.Unchecked)

        itemChild2 = QTreeWidgetItem(itemTop1)
        itemChild2.setFlags(itemChild2.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
        itemChild2.setText(0, '12_비유동자산')
        itemChild2.setCheckState(0, Qt.Unchecked)

        itemTop2 = QTreeWidgetItem(self.account_tree)
        itemTop2.setText(0, '2_Liability')
        itemTop2.setFlags(itemTop2.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)

        itemChild3 = QTreeWidgetItem(itemTop2)
        itemChild3.setFlags(itemChild3.flags() | Qt.ItemIsUserCheckable)
        itemChild3.setText(0, '21_유동부채')
        itemChild3.setCheckState(0, Qt.Unchecked)

        itemChild4 = QTreeWidgetItem(itemTop2)
        itemChild4.setFlags(itemChild4.flags() | Qt.ItemIsUserCheckable)
        itemChild4.setText(0, '22_비유동부채')
        itemChild4.setCheckState(0, Qt.Unchecked)

        labelJE = QLabel('전표입력자 : ', self.dialog8)
        labelJE.setStyleSheet("color: white;")

        font3 = labelJE.font()
        font3.setBold(True)
        labelJE.setFont(font3)

        self.D8_JE = QLineEdit(self.dialog8)
        self.D8_JE.setStyleSheet("background-color: white;")
        self.D8_JE.setPlaceholderText('전표입력자 ID를 입력하세요')

        labelCost = QLabel('중요성금액 : ', self.dialog8)
        labelCost.setStyleSheet("color: white;")

        font4 = labelCost.font()
        font4.setBold(True)
        labelCost.setFont(font4)

        self.D8_Cost = QLineEdit(self.dialog8)
        self.D8_Cost.setStyleSheet("background-color: white;")
        self.D8_Cost.setPlaceholderText('100,000,000원 이상 입력하세요')

        self.D8_N.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D8_JE.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D8_Cost.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        layout1 = QGridLayout()
        layout1.addWidget(labelDate, 0, 0)
        layout1.addWidget(self.D8_N, 0, 1)
        layout1.addWidget(label_tree, 1, 0)
        layout1.addWidget(self.account_tree, 1, 1)
        layout1.addWidget(labelJE, 2, 0)
        layout1.addWidget(self.D8_JE, 2, 1)
        layout1.addWidget(labelCost, 3, 0)
        layout1.addWidget(self.D8_Cost, 3, 1)

        layout2 = QHBoxLayout()
        layout2.addStretch()
        layout2.addStretch()
        layout2.addWidget(self.btn2)
        layout2.addWidget(self.btnDialog)

        layout2.setContentsMargins(-1, 10, -1, -1)

        main_layout = QVBoxLayout()
        main_layout.setAlignment(Qt.AlignTop)
        main_layout.addLayout(layout1)
        main_layout.addLayout(layout2)

        self.dialog8.setLayout(main_layout)
        self.dialog8.setGeometry(300, 300, 700, 400)

        self.dialog8.setWindowTitle("Scenario8")
        self.dialog8.setWindowModality(Qt.NonModal)
        self.dialog8.show()

    def Dialog9(self):
        self.dialog9 = QDialog()
        groupbox = QGroupBox('접속 정보')

        self.dialog9.setStyleSheet('background-color: #2E2E38')
        self.dialog9.setWindowIcon(QIcon("./EY_logo.png"))

        self.btn2 = QPushButton('   Extract Data', self.dialog9)
        self.btn2.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked9)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        self.btnDialog = QPushButton("  Close", self.dialog9)
        self.btnDialog.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close9)

        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)

        labelKeyword = QLabel('작성빈도수 : ', self.dialog9)
        labelKeyword.setStyleSheet("color: white;")

        font1 = labelKeyword.font()
        font1.setBold(True)
        labelKeyword.setFont(font1)

        self.D9_N = QLineEdit(self.dialog9)
        self.D9_N.setStyleSheet("background-color: white;")

        labelTE = QLabel('TE : ', self.dialog9)
        labelTE.setStyleSheet("color: white;")

        font2 = labelTE.font()
        font2.setBold(True)
        labelTE.setFont(font2)

        self.D9_TE = QLineEdit(self.dialog9)
        self.D9_TE.setStyleSheet("background-color: white;")

        self.btn2.resize(110, 30)
        self.btnDialog.resize(110, 30)

        self.D9_N.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D9_TE.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        layout1 = QGridLayout()
        layout1.addWidget(labelKeyword, 0, 0)
        layout1.addWidget(self.D9_N, 0, 1)
        layout1.addWidget(labelTE, 1, 0)
        layout1.addWidget(self.D9_TE, 1, 1)

        layout2 = QHBoxLayout()
        layout2.addStretch()
        layout2.addStretch()
        layout2.addWidget(self.btn2)
        layout2.addWidget(self.btnDialog)

        layout2.setContentsMargins(-1, 10, -1, -1)

        main_layout = QVBoxLayout()
        main_layout.setAlignment(Qt.AlignTop)
        main_layout.addLayout(layout1)
        main_layout.addLayout(layout2)

        # Extraction 내 Dictionary 를 위한 변수 설정
        self.D9_clickcount = 0

        self.dialog9.setLayout(main_layout)
        self.dialog9.setGeometry(300, 300, 500, 150)
        self.dialog9.setWindowTitle("Scenario9")
        self.dialog9.setWindowModality(Qt.NonModal)
        self.dialog9.show()

    def Dialog10(self):
        self.dialog10 = QDialog()
        self.dialog10.setStyleSheet('background-color: #2E2E38')
        self.dialog10.setWindowIcon(QIcon("./EY_logo.png"))

        self.btn2 = QPushButton('   Extract Data', self.dialog10)
        self.btn2.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked10)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        self.btnDialog = QPushButton("   Close", self.dialog10)
        self.btnDialog.setStyleSheet(
            'color:white;  background-image : url(./bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close10)

        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)

        self.btn2.resize(110, 30)
        self.btnDialog.resize(110, 30)

        labelKeyword = QLabel('전표입력자 : ', self.dialog10)
        labelKeyword.setStyleSheet("color: white;")

        font1 = labelKeyword.font()
        font1.setBold(True)
        labelKeyword.setFont(font1)

        self.D10_Search = QLineEdit(self.dialog10)
        self.D10_Search.setStyleSheet("background-color: white;")

        labelPoint = QLabel('특정시점 : ', self.dialog10)
        labelPoint.setStyleSheet("color: white;")

        font2 = labelPoint.font()
        font2.setBold(True)
        labelPoint.setFont(font2)

        self.D10_Point = QLineEdit(self.dialog10)
        self.D10_Point.setStyleSheet("background-color: white;")

        labelAccount = QLabel('특정계정 : ', self.dialog10)
        labelAccount.setStyleSheet("color: white;")

        font3 = labelAccount.font()
        font3.setBold(True)
        labelAccount.setFont(font3)

        self.D10_Account = QLineEdit(self.dialog10)
        self.D10_Account.setStyleSheet("background-color: white;")

        labelTE = QLabel('TE : ', self.dialog10)
        labelTE.setStyleSheet("color: white;")

        font4 = labelTE.font()
        font4.setBold(True)
        labelTE.setFont(font4)

        self.D10_TE = QLineEdit(self.dialog10)
        self.D10_TE.setStyleSheet("background-color: white;")

        self.D10_Search.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D10_Point.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D10_Account.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D10_TE.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        layout1 = QGridLayout()
        layout1.addWidget(labelKeyword, 0, 0)
        layout1.addWidget(self.D10_Search, 0, 1)
        layout1.addWidget(labelPoint, 1, 0)
        layout1.addWidget(self.D10_Point, 1, 1)
        layout1.addWidget(labelAccount, 2, 0)
        layout1.addWidget(self.D10_Account, 2, 1)
        layout1.addWidget(labelTE, 3, 0)
        layout1.addWidget(self.D10_TE, 3, 1)

        layout2 = QHBoxLayout()
        layout2.addStretch()
        layout2.addStretch()
        layout2.addWidget(self.btn2)
        layout2.addWidget(self.btnDialog)

        layout2.setContentsMargins(-1, 10, -1, -1)

        main_layout = QVBoxLayout()
        main_layout.setAlignment(Qt.AlignTop)
        main_layout.addLayout(layout1)
        main_layout.addLayout(layout2)

        self.dialog10.setLayout(main_layout)
        self.dialog10.setGeometry(300, 300, 500, 200)

        self.dialog10.setWindowTitle("Scenario10")
        self.dialog10.setWindowModality(Qt.NonModal)
        self.dialog10.show()

    # A,B,C조 작성 - 추후 논의
    # def Dialog11(self):

    def Dialog12(self):  # 수정완료
        self.dialog12 = QDialog()
        self.dialog12.setStyleSheet('background-color: #2E2E38')
        self.dialog12.setWindowIcon(QIcon('./EY_logo.png'))

        self.btn = QPushButton('   Extract Data', self.dialog12)
        self.btn.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btn.clicked.connect(self.extButtonClicked12)
        font9 = self.btn.font()
        font9.setBold(True)
        self.btn.setFont(font9)

        self.btnDialog = QPushButton("   Close", self.dialog12)
        self.btnDialog.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close12)
        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)
        self.btn.resize(110, 30)
        self.btnDialog.resize(110, 30)

        self.btn2 = QPushButton('   Extract Data', self.dialog12)
        self.btn2.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked12)
        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        self.btnDialog2 = QPushButton("   Close", self.dialog12)
        self.btnDialog2.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btnDialog2.clicked.connect(self.dialog_close12)
        font10 = self.btnDialog2.font()
        font10.setBold(True)
        self.btnDialog2.setFont(font10)
        self.btn2.resize(110, 30)
        self.btnDialog2.resize(110, 30)

        # 라벨값
        labelAccount = QLabel('Account Code* : ', self.dialog12)
        labelAccount.setStyleSheet("color: white;")
        font3 = labelAccount.font()
        font3.setBold(True)
        labelAccount.setFont(font3)
        labelCost = QLabel('중요성 금액 : ', self.dialog12)
        labelCost.setStyleSheet("color: white;")
        font3 = labelCost.font()
        font3.setBold(True)
        labelCost.setFont(font3)
        labelCost2 = QLabel('중요성 금액 : ', self.dialog12)
        labelCost2.setStyleSheet("color: white;")
        font4 = labelCost2.font()
        font4.setBold(True)
        labelCost2.setFont(font4)

        self.D12_Code = QTextEdit(self.dialog12)
        self.D12_Code.setAcceptRichText(False)
        self.D12_Code.setStyleSheet("background-color: white;")
        self.D12_Code.setPlaceholderText('계정코드를 입력하세요')

        self.D12_Cost = QLineEdit(self.dialog12)
        self.D12_Cost.setStyleSheet("background-color: white;")
        self.D12_Cost.setPlaceholderText('100,000,000원 이상 입력하세요')
        self.D12_Cost.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        self.D12_Cost2 = QLineEdit(self.dialog12)
        self.D12_Cost2.setStyleSheet("background-color: white;")
        self.D12_Cost2.setPlaceholderText('100,000,000원 이상 입력하세요')
        self.D12_Cost2.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        label_tree = QLabel('원하는 계정명을 선택하세요', self.dialog12)
        label_tree.setStyleSheet("color: white;")
        label_tree.setFont(QFont('Times, font', 9))
        font4 = label_tree.font()
        font4.setBold(True)
        label_tree.setFont(font4)

        # 트리 예시
        self.account_tree = QTreeWidget(self.dialog12)
        self.account_tree.setColumnCount(2)
        self.account_tree.setStyleSheet("background-color: white;")
        self.account_tree.setHeaderLabels(['Account Type'])
        self.account_tree.setAlternatingRowColors(False)
        self.account_tree.header().setVisible(True)

        itemTop1 = QTreeWidgetItem(self.account_tree)
        itemTop1.setText(0, "1_Assets")
        itemTop1.setFlags(itemTop1.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)

        itemChild1 = QTreeWidgetItem(itemTop1)
        itemChild1.setFlags(itemChild1.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
        itemChild1.setText(0, '11_유동자산')
        itemChild1.setCheckState(0, Qt.Unchecked)

        itemChild11 = QTreeWidgetItem(itemChild1)
        itemChild11.setFlags(itemChild11.flags() | Qt.ItemIsUserCheckable)
        itemChild11.setText(0, '1101_현금및현금성자산')
        itemChild11.setCheckState(0, Qt.Unchecked)

        itemChild12 = QTreeWidgetItem(itemChild1)
        itemChild12.setFlags(itemChild12.flags() | Qt.ItemIsUserCheckable)
        itemChild12.setText(0, '1105_매출채권')
        itemChild12.setCheckState(0, Qt.Unchecked)

        itemChild2 = QTreeWidgetItem(itemTop1)
        itemChild2.setFlags(itemChild2.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
        itemChild2.setText(0, '12_비유동자산')
        itemChild2.setCheckState(0, Qt.Unchecked)

        itemTop2 = QTreeWidgetItem(self.account_tree)
        itemTop2.setText(0, '2_Liability')
        itemTop2.setFlags(itemTop2.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)

        itemChild3 = QTreeWidgetItem(itemTop2)
        itemChild3.setFlags(itemChild3.flags() | Qt.ItemIsUserCheckable)
        itemChild3.setText(0, '21_유동부채')
        itemChild3.setCheckState(0, Qt.Unchecked)

        itemChild4 = QTreeWidgetItem(itemTop2)
        itemChild4.setFlags(itemChild4.flags() | Qt.ItemIsUserCheckable)
        itemChild4.setText(0, '22_비유동부채')
        itemChild4.setCheckState(0, Qt.Unchecked)

        tab1 = QWidget()
        tab2 = QWidget()
        tabs = QTabWidget()

        sublayout1 = QGridLayout()  # 계정 트리
        sublayout1.addWidget(label_tree, 0, 0)
        sublayout1.addWidget(self.account_tree, 1, 0)

        sublayout2 = QGridLayout()  # 계정코드 입력했을 때 - 텍스트에딧 추가
        sublayout2.addWidget(labelAccount, 0, 0)
        sublayout2.addWidget(self.D12_Code, 0, 1)
        sublayout2.addWidget(labelCost2, 1, 0)
        sublayout2.addWidget(self.D12_Cost2, 1, 1)

        sublayout3 = QGridLayout()  # 중요성 금액
        sublayout3.addWidget(labelCost, 0, 0)
        sublayout3.addWidget(self.D12_Cost, 0, 1)

        sublayout4 = QHBoxLayout()
        sublayout4.addStretch()
        sublayout4.addStretch()
        sublayout4.addWidget(self.btn)
        sublayout4.addWidget(self.btnDialog)

        sublayout5 = QHBoxLayout()
        sublayout5.addStretch()
        sublayout5.addStretch()
        sublayout5.addWidget(self.btn2)
        sublayout5.addWidget(self.btnDialog2)

        layout1 = QVBoxLayout()
        layout1.addLayout(sublayout1)
        layout1.addLayout(sublayout3)
        layout1.addLayout(sublayout4)

        layout2 = QVBoxLayout()
        layout2.addLayout(sublayout2)
        layout2.addLayout(sublayout5)

        main_layout = QVBoxLayout()
        tab1.setLayout(layout1)
        tab2.setLayout(layout2)
        tabs.addTab(tab1, "Account Name")
        tabs.addTab(tab2, "Account Code")
        main_layout.addWidget(tabs)

        self.dialog12.setLayout(main_layout)
        self.dialog12.resize(500, 400)
        self.dialog12.setWindowTitle('Scenario12')
        self.dialog12.setWindowModality(Qt.NonModal)
        self.dialog12.show()

    def Dialog13(self):
        self.dialog13 = QDialog()
        self.dialog13.setStyleSheet('background-color: #2E2E38')
        self.dialog13.setWindowIcon(QIcon('./EY_logo.png'))

        ### 버튼 - Extract Data
        self.btn2 = QPushButton(' Extract Data', self.dialog13)
        self.btn2.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked13)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        ### 버튼 - Close
        self.btnDialog = QPushButton('Close', self.dialog13)
        self.btnDialog.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close13)

        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)

        ### 버튼 3 - Save
        self.btnSave = QPushButton(' Save', self.dialog13)
        self.btnSave.setStyleSheet('color:white; background-image: url(./bar.png)')
        self.btnSave.clicked.connect(self.extButtonClicked13)

        font13 = self.btnSave.font()
        font13.setBold(True)
        self.btnSave.setFont(font13)

        ### 버튼 4 - Save and Process
        self.btnSaveProceed = QPushButton(' Save and Proceed', self.dialog13)
        self.btnSaveProceed.setStyleSheet('color: white; background-image: url(./bar.png)')
        self.btnSaveProceed.clicked.connect(self.extButtonClicked13)

        font14 = self.btnSaveProceed.font()
        font14.setBold(True)
        self.btnSaveProceed.setFont(font14)

        ### 라벨 1 - 연속된 자릿수
        label_Continuous = QLabel('연속된 자릿수(ex. 3333, 6666)*: ', self.dialog13)
        label_Continuous.setStyleSheet("color: red;")
        label_Continuous.setFont(QFont('Arial', 9))

        font1 = label_Continuous.font()
        font1.setBold(True)
        label_Continuous.setFont(font1)

        ### Text Edit - 연속된 자릿수
        self.text_continuous = QTextEdit(self.dialog13)
        self.text_continuous.setAcceptRichText(False)
        self.text_continuous.setStyleSheet("background-color: white;")

        ### 라벨 2 - 중요성 금액
        label_amount = QLabel('중요성금액 : ', self.dialog13)
        label_amount.setStyleSheet("color: white;")
        label_amount.setFont(QFont('Times font', 9))

        font3 = label_amount.font()
        font3.setBold(True)
        label_amount.setFont(font3)

        ### Line Edit - 중요성 금액
        self.line_amount = QLineEdit(self.dialog13)
        self.line_amount.setStyleSheet("background-color: white;")

        ### 라벨 3 - 계정 트리
        label_tree = QLabel('원하는 계정명을 선택하세요', self.dialog13)
        label_tree.setStyleSheet("color: white;")
        label_tree.setFont(QFont('Times, font', 9))

        font4 = label_tree.font()
        font4.setBold(True)
        label_tree.setFont(font4)

        ### TreeWidget
        self.account_tree = QTreeWidget(self.dialog13)
        self.account_tree.setColumnCount(2)
        self.account_tree.setStyleSheet("background-color: white;")
        self.account_tree.setHeaderLabels(['Account Type'])
        self.account_tree.setAlternatingRowColors(False)
        self.account_tree.header().setVisible(True)

        itemTop1 = QTreeWidgetItem(self.account_tree)
        itemTop1.setText(0, "1_Assets")
        itemTop1.setFlags(itemTop1.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)

        itemChild1 = QTreeWidgetItem(itemTop1)
        itemChild1.setFlags(itemChild1.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
        itemChild1.setText(0, '11_유동자산')
        itemChild1.setCheckState(0, Qt.Unchecked)

        itemChild11 = QTreeWidgetItem(itemChild1)
        itemChild11.setFlags(itemChild11.flags() | Qt.ItemIsUserCheckable)
        itemChild11.setText(0, '1101_현금및현금성자산')
        itemChild11.setCheckState(0, Qt.Unchecked)

        itemChild12 = QTreeWidgetItem(itemChild1)
        itemChild12.setFlags(itemChild12.flags() | Qt.ItemIsUserCheckable)
        itemChild12.setText(0, '1105_매출채권')
        itemChild12.setCheckState(0, Qt.Unchecked)

        itemChild2 = QTreeWidgetItem(itemTop1)
        itemChild2.setFlags(itemChild2.flags() | Qt.ItemIsUserCheckable)
        itemChild2.setText(0, '12_비유동자산')
        itemChild2.setCheckState(0, Qt.Unchecked)

        itemTop2 = QTreeWidgetItem(self.account_tree)
        itemTop2.setText(0, '2_Liability')
        itemTop2.setFlags(itemTop2.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)

        itemChild3 = QTreeWidgetItem(itemTop2)
        itemChild3.setFlags(itemChild3.flags() | Qt.ItemIsUserCheckable)
        itemChild3.setText(0, '21_유동부채')
        itemChild3.setCheckState(0, Qt.Unchecked)

        itemChild4 = QTreeWidgetItem(itemTop2)
        itemChild4.setFlags(itemChild4.flags() | Qt.ItemIsUserCheckable)
        itemChild4.setText(0, '22_비유동부채')
        itemChild4.setCheckState(0, Qt.Unchecked)

        ### Layout - 다이얼로그 UI
        main_layout = QVBoxLayout()

        layout1 = QVBoxLayout()
        sublayout1 = QVBoxLayout()
        sublayout2 = QHBoxLayout()

        layout2 = QVBoxLayout()
        sublayout3 = QVBoxLayout()
        sublayout4 = QHBoxLayout()

        tab1 = QWidget()
        tab2 = QWidget()
        tabs = QTabWidget()

        sublayout1.addWidget(label_Continuous)
        sublayout1.addWidget(self.text_continuous)
        sublayout1.addWidget(label_amount)
        sublayout1.addWidget(self.line_amount)

        sublayout2.addStretch(1)
        sublayout2.addWidget(self.btnSave, stretch=1, alignment=Qt.AlignBottom)
        sublayout2.addWidget(self.btnSaveProceed, stretch=1, alignment=Qt.AlignBottom)
        sublayout2.addStretch(1)

        layout1.addLayout(sublayout1, stretch=4)
        layout1.addLayout(sublayout2, stretch=1)

        sublayout3.addWidget(label_tree)
        sublayout3.addWidget(self.account_tree)

        sublayout4.addStretch(1)
        sublayout4.addWidget(self.btn2, stretch=1, alignment=Qt.AlignBottom)
        sublayout4.addWidget(self.btnDialog, stretch=1, alignment=Qt.AlignBottom)
        sublayout4.addStretch(1)

        layout2.addLayout(sublayout3, stretch=4)
        layout2.addLayout(sublayout4, stretch=1)

        tab1.setLayout(layout1)
        tab2.setLayout(layout2)

        tabs.addTab(tab1, "자릿수/금액")
        tabs.addTab(tab2, "계정 선택")

        main_layout.addWidget(tabs)

        self.dialog13.setLayout(main_layout)

        self.dialog13.resize(500, 400)

        self.dialog13.setWindowTitle('Scenario13')
        self.dialog13.setWindowModality(Qt.NonModal)
        self.dialog13.show()

    def Dialog14(self):
        self.dialog14 = QDialog()
        self.dialog14.setStyleSheet('background-color: #2E2E38')
        self.dialog14.setWindowIcon(QIcon("./EY_logo.png"))

        self.btn2 = QPushButton('   Extract Data', self.dialog14)
        self.btn2.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked14)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        self.btnDialog = QPushButton("   Close", self.dialog14)
        self.btnDialog.setStyleSheet(
            'color:white;  background-image : url(./bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close14)

        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)

        self.btn2.resize(110, 30)
        self.btnDialog.resize(110, 30)

        labelKeyword = QLabel('Key Words : ', self.dialog14)
        labelKeyword.setStyleSheet("color: white;")

        font1 = labelKeyword.font()
        font1.setBold(True)
        labelKeyword.setFont(font1)

        self.D14_Key = QLineEdit(self.dialog14)
        self.D14_Key.setStyleSheet("background-color: white;")

        labelTE = QLabel('TE : ', self.dialog14)
        labelTE.setStyleSheet("color: white;")

        font2 = labelTE.font()
        font2.setBold(True)
        labelTE.setFont(font2)

        self.D14_TE = QLineEdit(self.dialog14)
        self.D14_TE.setStyleSheet("background-color: white;")

        self.D14_Key.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D14_TE.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        layout1 = QGridLayout()
        layout1.addWidget(labelKeyword, 0, 0)
        layout1.addWidget(self.D14_Key, 0, 1)
        layout1.addWidget(labelTE, 1, 0)
        layout1.addWidget(self.D14_TE, 1, 1)

        layout2 = QHBoxLayout()
        layout2.addStretch()
        layout2.addStretch()
        layout2.addWidget(self.btn2)
        layout2.addWidget(self.btnDialog)

        layout2.setContentsMargins(-1, 10, -1, -1)

        main_layout = QVBoxLayout()
        main_layout.setAlignment(Qt.AlignTop)
        main_layout.addLayout(layout1)
        main_layout.addLayout(layout2)

        self.dialog14.setLayout(main_layout)
        self.dialog14.setGeometry(300, 300, 500, 150)

        self.dialog14.setWindowTitle("Scenario14")
        self.dialog14.setWindowModality(Qt.NonModal)
        self.dialog14.show()

    def dialog_close4(self):
        self.dialog4.close()

    def dialog_close5(self):
        self.dialog5.close()

    def dialog_close6(self):
        self.dialog6.close()

    def dialog_close7(self):
        self.dialog7.close()

    def dialog_close8(self):
        self.dialog8.close()

    def dialog_close9(self):
        self.dialog9.close()

    def dialog_close10(self):
        self.dialog10.close()

    def dialog_close11(self):
        self.dialog11.close()

    def dialog_close12(self):
        self.dialog12.close()

    def dialog_close13(self):
        self.dialog13.close()

    def dialog_close14(self):
        self.dialog14.close()

    def Show_DataFrame_Group(self):
        tables = QGroupBox('데이터')
        self.setStyleSheet('QGroupBox  {color: white;}')
        font6 = tables.font()
        font6.setBold(True)
        tables.setFont(font6)
        box = QBoxLayout(QBoxLayout.TopToBottom)

        self.viewtable = QTableView(self)

        box.addWidget(self.viewtable)
        tables.setLayout(box)

        return tables

    ###########################
    ####추가하셔야 하는 함수들#####
    ###########################
    def Sheet_ComboBox_Selected(self, text):
        model = DataFrameModel(self.scenario_dic[text])
        self.viewtable.setModel(model)
        self.selected_scenario_group = text

    def RemoveSheetButton_Clicked(self):
        temp = self.combo_sheet.findText(self.selected_scenario_group)
        self.combo_sheet.removeItem(temp)
        del self.scenario_dic[self.selected_scenario_group]

    def Save_Buttons_Group(self):
        ##GroupBox
        groupbox = QGroupBox("저장")
        font_groupbox = groupbox.font()
        font_groupbox.setBold(True)
        groupbox.setFont(font_groupbox)
        self.setStyleSheet('QGroupBox  {color: white;}')

        ##GroupBox에 넣을 Layout들
        layout = QHBoxLayout()
        left_sublayout = QGridLayout()
        right_sublayout1 = QVBoxLayout()
        right_sublayout2 = QHBoxLayout()
        right_sublayout3 = QHBoxLayout()

        ##RemoveSheet 버튼
        RemoveSheet_button = QPushButton('Remove Sheet')
        RemoveSheet_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        RemoveSheet_button.setStyleSheet('color:white;background-image : url(./bar.png)')
        font_RemoveSheet = RemoveSheet_button.font()
        font_RemoveSheet.setBold(True)
        RemoveSheet_button.setFont(font_RemoveSheet)

        # label
        label_sheet = QLabel("Sheet names: ", self)
        font_sheet = label_sheet.font()
        font_sheet.setBold(True)
        label_sheet.setFont(font_sheet)
        label_sheet.setStyleSheet('color:white;')

        label_savepath = QLabel(f"Save Route: {' ' * 12}", self)
        font_savepath = label_savepath.font()
        font_savepath.setBold(True)
        label_savepath.setFont(font_savepath)
        label_savepath.setStyleSheet('color:white;')

        ##시나리오 Sheet를 표현할 콤보박스
        self.combo_sheet = QComboBox(self)

        ##저장 경로를 표현할 LineEdit
        self.line_savepath = QLineEdit(self)
        self.line_savepath.setText("")
        self.line_savepath.setDisabled(True)

        ## Setting Save Route 버튼
        save_path_button = QPushButton("Setting Save Route", self)
        save_path_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        font_path_button = save_path_button.font()
        font_path_button.setBold(True)
        save_path_button.setFont(font_path_button)
        save_path_button.setStyleSheet('color:white;background-image : url(./bar.png)')

        ## Save 버튼
        export_file_button = QPushButton("Save", self)
        export_file_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        font_export_button = export_file_button.font()
        font_export_button.setBold(True)
        export_file_button.setFont(font_export_button)
        export_file_button.setStyleSheet('color:white;background-image : url(./bar.png)')

        #########
        #########버튼 클릭 or 콤보박스 선택시 발생하는 시그널 함수들
        RemoveSheet_button.clicked.connect(self.RemoveSheetButton_Clicked)
        save_path_button.clicked.connect(self.saveFileDialog)
        export_file_button.clicked.connect(self.saveFile)
        self.combo_sheet.activated[str].connect(self.Sheet_ComboBox_Selected)

        ##layout 쌓기
        left_sublayout.addWidget(label_sheet, 0, 0)
        left_sublayout.addWidget(self.combo_sheet, 0, 1)
        left_sublayout.addWidget(label_savepath, 1, 0)
        left_sublayout.addWidget(self.line_savepath, 1, 1)

        right_sublayout2.addWidget(RemoveSheet_button, stretch=2)
        right_sublayout3.addWidget(save_path_button, stretch=1)
        right_sublayout3.addWidget(export_file_button, stretch=1)

        right_sublayout1.addLayout(right_sublayout2, stretch=1)
        right_sublayout1.addLayout(right_sublayout3, stretch=1)

        layout.addLayout(left_sublayout, stretch=2)
        layout.addLayout(right_sublayout1, stretch=1)

        groupbox.setLayout(layout)

        return groupbox

    def handle_date_clicked(self, date):
        self.D6_Date.setText(date.toString("yyyy-MM-dd"))

    def handle_date_clicked2(self, date):
        self.D7_Date.setText(date.toString("yyyy-MM-dd"))

    def calendar(self):
        self.new_calendar.show()

    def extButtonClicked4(self):

        temp_N = self.D4_N.text()
        temp_TE = self.D4_TE.text()

        if temp_N == '':
            self.alertbox_open()

        else:
            cursor = self.cnxn.cursor()

            sql_query = """
            """.format(field=self.selected_project_id)

        self.dataframe = pd.read_sql(sql_query, self.cnxn)

        model = DataFrameModel(self.dataframe)
        self.viewtable.setModel(model)

    def extButtonClicked5_SAP(self):

        ### ListBox 인풋값 append
        dropped_items = []
        for i in range(self.listbox_drops.count()):
            dropped_items.append(self.listbox_drops.item(i))

        ### 파일 경로 unicode 문제 해결
        for i in range(self.dropped_items.count()):
            dropped_items[i] = re.sub(r'\'', '/', dropped_items[i])

        ### dataframe으로 저장
        df = pd.DataFrame()
        for i in range(len(dropped_items)):
            df = df.append(pd.read_csv(dropped_items[i], sep='|'))

        ### 당기 생성된 계정 코드 반환
        temp_AccCode = list()

        for i in range(len(df)):
            df.loc[i, 'ERDAT'] = str(df.loc[i, 'ERDAT'])
            year = df.loc[i, 'ERDAT'][0:4]

            ### 당기 시점 지정
            now = datetime.datetime.now()
            before_three_months = now - relativedelta(month=3)

            if int(year) == before_three_months.year:
                temp_AccCode.append(df.loc[i, 'SAKNR'])

        if temp_AccCode == '':
            self.alertbox_open()

        else:
            cursor = self.cnxn.cursor()

            sql_query = """""".format(field=self.selected_project_id)

        self.dataframe = pd.read_sql(sql_query, self.cnxn)

        model = DataFrameModel(self.dataframe)
        self.viewtable.setModel(model)

    def extButtonClicked5_Non_SAP(self):

        temp_Code_Non_SAP = self.D5_Code.text()
        temp_Code_Non_SAP = re.sub(r"[:,|\s]", ",", temp_Code_Non_SAP)
        temp_Code_Non_SAP = re.split(",", temp_Code_Non_SAP)

        if temp_Code_Non_SAP == '':
            self.alertbox_open()

        else:
            cursor = self.cnxn.cursor()

            sql_query = """""".format(field=self.selected_project_id)

        self.dataframe = pd.read_sql(sql_query, self.cnxn)

        model = DataFrameModel(self.dataframe)
        self.viewtable.setModel(model)

    def extButtonClicked6(self):

        tempDate = self.D6_Date.text()  # 필수값
        tempTDate = self.D6_Date2.text()  # 필수값
        tempAccount = self.D6_Account.text()
        tempJE = self.D6_JE.text()
        tempCost = self.D6_Cost.text()

        if tempTDate == '' or tempDate == '':
            self.alertbox_open()

        else:
            if tempCost == '': tempCost = 0

            try:
                int(tempTDate)
                int(tempCost)
                cursor = self.cnxn.cursor()

                # sql문 수정
                sql = '''
                               SELECT TOP 100											
                                   JournalEntries.BusinessUnit											
                                   , JournalEntries.JENumber											
                                   , JournalEntries.JELineNumber											
                                   , JournalEntries.EffectiveDate											
                                   , JournalEntries.EntryDate											
                                   , JournalEntries.Period											
                                   , JournalEntries.GLAccountNumber											
                                   , CoA.GLAccountName											
                                   , JournalEntries.Debit											
                                   , JournalEntries.Credit											
                                   , CASE
                                        WHEN JournalEntries.Debit = 0 THEN 'Credit' ELSE 'Debit'
                                        END AS DebitCredit
                                   , JournalEntries.Amount											
                                   , JournalEntries.FunctionalCurrencyCode											
                                   , JournalEntries.JEDescription											
                                   , JournalEntries.JELineDescription											
                                   , JournalEntries.Source											
                                   , JournalEntries.PreparerID											
                                   , JournalEntries.ApproverID											
                               FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] JournalEntries,											
                                       [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											
                               WHERE JournalEntries.GLAccountNumber = CoA.GLAccountNumber 
                                ORDER BY JENumber, JELineNumber											
                            '''.format(field=self.selected_project_id)

                self.dataframe = pd.read_sql(sql, self.cnxn)

                model = DataFrameModel(self.dataframe)
                self.viewtable.setModel(model)

            except ValueError:
                try:
                    int(tempTDate)
                    try:
                        int(tempCost)
                    except:
                        self.alertbox_open2('중요성금액')
                except:
                    try:
                        int(tempCost)
                        self.alertbox_open2('T')
                    except:
                        self.alertbox_open2('T값과 중요성금액')

    def extButtonClicked7(self):

        tempDate = self.D7_Date.text()  # 필수값
        tempAccount = self.D7_Account.text()
        tempJE = self.D7_JE.text()
        tempCost = self.D7_Cost.text()

        if self.rbtn1.isChecked():
            tempState = 'Effective Date'

        elif self.rbtn2.isChecked():
            tempState = 'Entry Date'

        if tempCost == '':
            tempCost = 0

        if tempDate == '':
            self.alertbox_open()

        else:
            try:
                int(tempCost)
                cursor = self.cnxn.cursor()

                # sql문 수정
                sql = '''
                               SELECT TOP 100											
                                   JournalEntries.BusinessUnit											
                                   , JournalEntries.JENumber											
                                   , JournalEntries.JELineNumber											
                                   , JournalEntries.EffectiveDate											
                                   , JournalEntries.EntryDate											
                                   , JournalEntries.Period											
                                   , JournalEntries.GLAccountNumber											
                                   , CoA.GLAccountName											
                                   , JournalEntries.Debit											
                                   , JournalEntries.Credit											
                                   , CASE
                                       WHEN JournalEntries.Debit = 0 THEN 'Credit' ELSE 'Debit'
                                       END AS DebitCredit
                                   , JournalEntries.Amount											
                                   , JournalEntries.FunctionalCurrencyCode											
                                   , JournalEntries.JEDescription											
                                   , JournalEntries.JELineDescription											
                                   , JournalEntries.Source											
                                   , JournalEntries.PreparerID											
                                   , JournalEntries.ApproverID											
                               FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] JournalEntries,											
                                       [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											
                               WHERE JournalEntries.GLAccountNumber = CoA.GLAccountNumber 
                               ORDER BY JENumber, JELineNumber											
                           '''.format(field=self.selected_project_id)

                self.dataframe = pd.read_sql(sql, self.cnxn)

                model = DataFrameModel(self.dataframe)
                self.viewtable.setModel(model)

            except ValueError:
                self.alertbox_open2('중요성 금액')

    def extButtonClicked8(self):

        tempN = self.D8_N.text()  # 필수값
        tempAccount = self.D8_Account.text()
        tempJE = self.D8_JE.text()
        tempCost = self.D8_Cost.text()

        if tempN == '':
            self.alertbox_open()

        else:
            if tempCost == '': tempCost = 0
            try:
                int(tempN)
                int(tempCost)
                cursor = self.cnxn.cursor()

                # sql문 수정
                sql = '''
                               SELECT TOP 100											
                                   JournalEntries.BusinessUnit											
                                   , JournalEntries.JENumber											
                                   , JournalEntries.JELineNumber											
                                   , JournalEntries.EffectiveDate											
                                   , JournalEntries.EntryDate											
                                   , JournalEntries.Period											
                                   , JournalEntries.GLAccountNumber											
                                   , CoA.GLAccountName											
                                   , JournalEntries.Debit											
                                   , JournalEntries.Credit											
                                   , CASE
                                        WHEN JournalEntries.Debit = 0 THEN 'Credit' ELSE 'Debit'
                                        END AS DebitCredit
                                   , JournalEntries.Amount											
                                   , JournalEntries.FunctionalCurrencyCode											
                                   , JournalEntries.JEDescription											
                                   , JournalEntries.JELineDescription											
                                   , JournalEntries.Source											
                                   , JournalEntries.PreparerID											
                                   , JournalEntries.ApproverID											
                               FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] JournalEntries,											
                                       [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											
                               WHERE JournalEntries.GLAccountNumber = CoA.GLAccountNumber 
                               ORDER BY JENumber, JELineNumber											
                            '''.format(field=self.selected_project_id)

                self.dataframe = pd.read_sql(sql, self.cnxn)

                model = DataFrameModel(self.dataframe)
                self.viewtable.setModel(model)

            except ValueError:
                try:
                    int(tempN)
                    try:
                        int(tempCost)
                    except:
                        self.alertbox_open2('중요성금액')
                except:
                    try:
                        int(tempCost)
                        self.alertbox_open2('N')
                    except:
                        self.alertbox_open2('N값과 중요성금액')

    def extButtonClicked9(self):
        # 다이얼로그별 Clickcount 설정
        self.D9_clickcount = self.D9_clickcount + 1
        tempN = self.D9_N.text()  # 필수값
        tempTE = self.D9_TE.text()

        if tempN == '' or tempTE == '':
            self.alertbox_open()

        else:
            cursor = self.cnxn.cursor()

            # sql문 수정
            sql = '''
                   SELECT TOP 100											
                       JournalEntries.BusinessUnit											
                       , JournalEntries.JENumber											
                       , JournalEntries.JELineNumber											
                       , JournalEntries.EffectiveDate											
                       , JournalEntries.EntryDate											
                       , JournalEntries.Period											
                       , JournalEntries.GLAccountNumber											
                       , CoA.GLAccountName											
                       , JournalEntries.Debit											
                       , JournalEntries.Credit											
                       , CASE
                            WHEN JournalEntries.Debit = 0 THEN 'Credit' ELSE 'Debit'
                            END AS DebitCredit
                       , JournalEntries.Amount											
                       , JournalEntries.FunctionalCurrencyCode											
                       , JournalEntries.JEDescription											
                       , JournalEntries.JELineDescription											
                       , JournalEntries.Source											
                       , JournalEntries.PreparerID											
                       , JournalEntries.ApproverID											
                   FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] JournalEntries,											
                           [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											
                   WHERE JournalEntries.GLAccountNumber = CoA.GLAccountNumber AND JournalEntries.Amount > {amount}
                   ORDER BY JENumber, JELineNumber											
                '''.format(field=self.selected_project_id, amount=tempTE)

            self.dataframe = pd.read_sql(sql, self.cnxn)

            # 딕셔너리 선언 및 시나리오 콤보 박스 추가
            self.scenario_dic['전표 작성 빈도수가 N회 이하 Scenario_' + str(
                self.D9_clickcount) + ' (N = ' + tempN + ', TE = ' + tempTE + ')'] = self.dataframe
            key_list = list(self.scenario_dic.keys())
            result = [key_list[0], key_list[-1]]
            model = DataFrameModel(self.scenario_dic[result[1]])
            self.combo_sheet.addItem(str(result[1]))
            self.viewtable.setModel(model)

    def extButtonClicked10(self):

        tempSearch = self.D10_Search.text()  # 필수값
        tempAccount = self.D10_Account.text()
        tempPoint = self.D10_Point.text()
        tempTE = self.D10_TE.text()

        if tempSearch == '' or tempAccount == '' or tempPoint == '' or tempTE == '':
            self.alertbox_open()

        else:
            cursor = self.cnxn.cursor()

            # sql문 수정
            sql = '''
                   SELECT TOP 100											
                       JournalEntries.BusinessUnit											
                       , JournalEntries.JENumber											
                       , JournalEntries.JELineNumber											
                       , JournalEntries.EffectiveDate											
                       , JournalEntries.EntryDate											
                       , JournalEntries.Period											
                       , JournalEntries.GLAccountNumber											
                       , CoA.GLAccountName											
                       , JournalEntries.Debit											
                       , JournalEntries.Credit											
                       , CASE
                            WHEN JournalEntries.Debit = 0 THEN 'Credit' ELSE 'Debit'
                            END AS DebitCredit
                       , JournalEntries.Amount											
                       , JournalEntries.FunctionalCurrencyCode											
                       , JournalEntries.JEDescription											
                       , JournalEntries.JELineDescription											
                       , JournalEntries.Source											
                       , JournalEntries.PreparerID											
                       , JournalEntries.ApproverID											
                   FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] JournalEntries,											
                           [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											
                   WHERE JournalEntries.GLAccountNumber = CoA.GLAccountNumber 
                   ORDER BY JENumber, JELineNumber											
                '''.format(field=self.selected_project_id)

            self.dataframe = pd.read_sql(sql, self.cnxn)

            model = DataFrameModel(self.dataframe)
            self.viewtable.setModel(model)

    def extButtonClicked12(self):

        tempCode = self.D12_Code.text()
        tempCost = self.D12_Cost.text()

        if self.rbtn1.isChecked():
            tempState = 'Account Name'

        elif self.rbtn2.isChecked():
            tempState = 'Account Code'

        if tempCost == '':
            tempCost = 0

        if tempState == '':
            self.alertbox_open()

        else:
            try:
                int(tempCost)
                cursor = self.cnxn.cursor()

                # sql문 수정
                sql = '''
                               SELECT TOP 100											
                                   JournalEntries.BusinessUnit											
                                   , JournalEntries.JENumber											
                                   , JournalEntries.JELineNumber											
                                   , JournalEntries.EffectiveDate											
                                   , JournalEntries.EntryDate											
                                   , JournalEntries.Period											
                                   , JournalEntries.GLAccountNumber											
                                   , CoA.GLAccountName											
                                   , JournalEntries.Debit											
                                   , JournalEntries.Credit											
                                   , CASE
                                       WHEN JournalEntries.Debit = 0 THEN 'Credit' ELSE 'Debit'
                                       END AS DebitCredit
                                   , JournalEntries.Amount											
                                   , JournalEntries.FunctionalCurrencyCode											
                                   , JournalEntries.JEDescription											
                                   , JournalEntries.JELineDescription											
                                   , JournalEntries.Source											
                                   , JournalEntries.PreparerID											
                                   , JournalEntries.ApproverID											
                               FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] JournalEntries,											
                                       [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											
                               WHERE JournalEntries.GLAccountNumber = CoA.GLAccountNumber 
                               ORDER BY JENumber, JELineNumber											
                           '''.format(field=self.selected_project_id)

                self.dataframe = pd.read_sql(sql, self.cnxn)

                model = DataFrameModel(self.dataframe)
                self.viewtable.setModel(model)

            except ValueError:
                self.alertbox_open2('중요성 금액')

    def extButtonClicked13(self):

        temp_Continuous = self.text_continuous.text()  # 필수
        temp_Tree = self.account_tree.text()
        temp_TE_13 = self.line_amount.text()

        if temp_Continuous == '' or temp_TE_13 == '' or temp_Tree == '':
            self.alertbox_open()

        else:
            cursor = self.cnxn.cursor()

            # sql문 수정
            sql_query = '''
            '''.format(field=self.selected_project_id)

        self.dataframe = pd.read_sql(sql_query, self.cnxn)
        model = DataFrameModel(self.dataframe)
        self.viewtable.setModel(model)

    def extButtonClicked14(self):

        tempKey = self.D14_Key.text()  # 필수값
        tempTE = self.D14_TE.text()

        if tempKey == '' or tempTE == '':
            self.alertbox_open()

        else:
            cursor = self.cnxn.cursor()

            # sql문 수정
            sql = '''
                   SELECT TOP 100											
                       JournalEntries.BusinessUnit											
                       , JournalEntries.JENumber											
                       , JournalEntries.JELineNumber											
                       , JournalEntries.EffectiveDate											
                       , JournalEntries.EntryDate											
                       , JournalEntries.Period											
                       , JournalEntries.GLAccountNumber											
                       , CoA.GLAccountName											
                       , JournalEntries.Debit											
                       , JournalEntries.Credit											
                       , CASE
                            WHEN JournalEntries.Debit = 0 THEN 'Credit' ELSE 'Debit'
                            END AS DebitCredit
                       , JournalEntries.Amount											
                       , JournalEntries.FunctionalCurrencyCode											
                       , JournalEntries.JEDescription											
                       , JournalEntries.JELineDescription											
                       , JournalEntries.Source											
                       , JournalEntries.PreparerID											
                       , JournalEntries.ApproverID											
                   FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] JournalEntries,											
                           [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											
                   WHERE JournalEntries.GLAccountNumber = CoA.GLAccountNumber 
                   ORDER BY JENumber, JELineNumber											
                '''.format(field=self.selected_project_id)

            self.dataframe = pd.read_sql(sql, self.cnxn)

            model = DataFrameModel(self.dataframe)
            self.viewtable.setModel(model)


    @pyqtSlot(QModelIndex)
    def slot_clicked_item(self, QModelIndex):
        self.stk_w.setCurrentIndex(QModelIndex.row())

    def saveFileDialog(self):
        fileName = QFileDialog.getSaveFileName(self, "Save File", '', ".xlsx")

        if fileName[0]:
            self.SaveRoute = fileName[0] + fileName[1]
            self.line_savepath.setText(self.SaveRoute)
        else:
            self.MessageBox_Open("저장 경로를 선택하지 않았습니다.")


    def saveFile(self):
        if self.dataframe is None:
            self.MessageBox_Open("저장할 데이터가 없습니다")
            return


        if self.scenario_dic == {}:
            self.MessageBox_Open("저장할 Sheet가 없습니다")
            return


        if self.SaveRoute == '' or self.SaveRoute is None:
            self.MessageBox_Open("저장 경로가 지정되지 않았습니다")
            return


        with pd.ExcelWriter(self.SaveRoute, engine='xlsxwriter') as writer:

            for sheet_name, df in self.scenario_dic.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False, encoding='utf-8')



if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())
