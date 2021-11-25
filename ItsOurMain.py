import sys
import re
import datetime
import time
from datetime import date
from datetime import timedelta
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

        self.setGeometry(1050, 400, 400, 200)
        self.setWindowTitle("PyQt5 QCalendar")
        self.setWindowIcon(QIcon("python.png"))
        self.setWindowModality(Qt.NonModal)

        vbox = QVBoxLayout()
        self.calendar = QCalendarWidget()
        self.calendar.setGridVisible(True)

        self.label = QLabel("")
        self.label.setStyleSheet('color:red')

        vbox.addWidget(self.calendar)
        vbox.addWidget(self.label)

        self.setLayout(vbox)


class Form(QGroupBox):
    def __init__(self, parent):
        super(Form, self).__init__(parent)

        grid = QGridLayout()
        self.setLayout(grid)
        self.setStyleSheet('QGroupBox  {color: white; background-color: white}')

        self.tree = QTreeWidget(self)
        self.tree.setStyleSheet("border-style: outset; border-color : white; background-color:white;")

        headerItem = QTreeWidgetItem()
        item = QTreeWidgetItem()

        grid.addWidget(self.tree, 0, 0)
        self.tree.setHeaderHidden(True)
        self.tree.itemClicked.connect(self.get_selected_leaves)

    def get_selected_leaves(self):
        checked_items = []

        def recurse(parent):
            for i in range(parent.childCount()):
                child = parent.child(i)
                for j in range(child.childCount()):
                    grandchild = child.child(j)
                    grandgrandchild = grandchild.childCount()
                    if grandgrandchild > 0:
                        recurse(grandchild)
                    else:
                        if grandchild.checkState(0) == Qt.Checked:
                            checked_items.append(grandchild.text(0).split(' ')[0])

        recurse(self.tree.invisibleRootItem())

        checked_name = ''
        for i in checked_items:
            checked_name = checked_name + ',' + '\'' + i + '\''

        checked_name = checked_name[1:]

        global checked_account

        checked_account = 'AND JournalEntries.GLAccountNumber IN (' + checked_name + ')'


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
        self.new_tree = None

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

    def alertbox_open3(self):
        self.alt = QMessageBox()
        self.alt.setIcon(QMessageBox.Information)
        self.alt.setWindowTitle('최대 라인 수 초과 오류')
        self.alt.setWindowIcon(QIcon("./EY_logo.png"))
        self.alt.setText('최대 라인 수가 초과 되었습니다.')
        self.alt.exec_()

    def alertbox_open4(self, state):
        self.alt = QMessageBox()
        self.alt.setIcon(QMessageBox.Information)
        txt = state
        self.alt.setWindowTitle('입력값 타입 오류')
        self.alt.setWindowIcon(QIcon("./EY_logo.png"))
        self.alt.setText(txt + '값을 8자리의 숫자로 입력해주시길 바랍니다.')
        self.alt.exec_()

    def alertbox_open5(self):
        self.alt = QMessageBox()
        self.alt.setIcon(QMessageBox.Information)
        self.alt.setWindowTitle('시트명 중복')
        self.alt.setWindowIcon(QIcon("./EY_logo.png"))
        self.alt.setText('이미 해당 시트명이 존재합니다.')
        self.alt.exec_()

    def alertbox_open6(self):
        self.alt = QMessageBox()
        self.alt.setIcon(QMessageBox.Information)
        self.alt.setWindowTitle('계정 선택 오류')
        self.alt.setWindowIcon(QIcon("./EY_logo.png"))
        self.alt.setText('계정이 선택되어 있지 않습니다.')
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
        ecode = self.line_ecode.text().strip()  ##leading/trailing space 포함시 제거
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

        ### Scenario 유형 콤보박스 - 소분류 수정
        self.comboScenario = QComboBox(self)

        self.comboScenario.addItem('--시나리오 목록--', [''])
        self.comboScenario.addItem('04 : 계정 사용빈도 N번 이하인 계정이 포함된 전표리스트', [''])
        self.comboScenario.addItem('05 : 당기 생성된 계정리스트 추출', [''])
        self.comboScenario.addItem('06 : 결산일 전후 T일 입력 전표', [''])
        self.comboScenario.addItem('07 : 비영업일 전기/입력 전표', [''])
        self.comboScenario.addItem('08 : 효력, 입력 일자 간 차이가 N일 이상인 전표', [''])
        self.comboScenario.addItem('09 : 전표 작성 빈도수가 N회 이하인 작성자에 의한 생성된 전표', [''])
        self.comboScenario.addItem('10 : 특정 전표 입력자(W)에 의해 생성된 전표', [''])
        self.comboScenario.addItem('11 : 특정한 주계정(A)과 특정한 상대계정(B)이 아닌 전표리스트 검토', [''])
        self.comboScenario.addItem('12 : 특정 계정(A)이 감소할 때 상대계정 리스트 검토', [''])
        self.comboScenario.addItem('13 : 연속된 숫자로 끝나는 금액 검토', [''])
        self.comboScenario.addItem('14 : 전표 description에 공란 또는 특정단어(key word)가 입력되어 있는 전표 리스트 (TE금액 제시 가능)', [''])

        self.ProjectCombobox = QComboBox(self)

        ##Engagement code 입력 line
        self.line_ecode = QLineEdit(self)
        self.line_ecode.setText("")

        ##Project Connect 버튼 생성 및 스타일 지정
        btn_connect = QPushButton('   Project Connect', self)
        font_btn_connect = btn_connect.font()
        font_btn_connect.setBold(True)
        btn_connect.setFont(font_btn_connect)
        btn_connect.setStyleSheet('color:white;  background-image : url(./bar.png)')

        ##Input Conditions 버튼 생성 및 스타일 지정
        btn_condition = QPushButton('   Input Conditions', self)
        font_btn_condition = btn_condition.font()
        font_btn_condition.setBold(True)
        btn_condition.setFont(font_btn_condition)
        btn_condition.setStyleSheet('color:white;  background-image : url(./bar.png)')

        ### Signal 함수들
        self.comboScenario.activated[str].connect(self.ComboSmall_Selected)
        self.cb_server.activated[str].connect(self.Server_ComboBox_Selected)
        btn_connect.clicked.connect(self.connectButtonClicked)
        self.ProjectCombobox.activated[str].connect(self.Project_ComboBox_Selected)
        btn_condition.clicked.connect(self.connectDialog)

        ##layout 쌓기
        grid = QGridLayout()
        grid.addWidget(label1, 0, 0)
        grid.addWidget(label2, 1, 0)
        grid.addWidget(label3, 2, 0)
        grid.addWidget(label4, 3, 0)
        grid.addWidget(self.cb_server, 0, 1)
        grid.addWidget(btn_connect, 1, 2)
        grid.addWidget(self.comboScenario, 3, 1)
        grid.addWidget(btn_condition, 3, 2)
        grid.addWidget(self.line_ecode, 1, 1)
        grid.addWidget(self.ProjectCombobox, 2, 1)

        groupbox.setLayout(grid)
        return groupbox

    def ComboSmall_Selected(self, text):
        self.selected_scenario_subclass_index = self.comboScenario.currentIndex()

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

        elif self.selected_scenario_class_index == 0 and self.selected_scenario_subclass_index == 3:
            self.Dialog6()

        elif self.selected_scenario_class_index == 0 and self.selected_scenario_subclass_index == 4:
            self.Dialog7()

        elif self.selected_scenario_class_index == 0 and self.selected_scenario_subclass_index == 5:
            self.Dialog8()

        elif self.selected_scenario_class_index == 0 and self.selected_scenario_subclass_index == 6:
            self.Dialog9()

        elif self.selected_scenario_class_index == 0 and self.selected_scenario_subclass_index == 7:
            self.Dialog10()

        elif self.selected_scenario_class_index == 0 and self.selected_scenario_subclass_index == 8:
            self.Dialog11()

        elif self.selected_scenario_class_index == 0 and self.selected_scenario_subclass_index == 9:
            self.Dialog12()

        elif self.selected_scenario_class_index == 0 and self.selected_scenario_subclass_index == 10:
            self.Dialog13()

        elif self.selected_scenario_class_index == 0 and self.selected_scenario_subclass_index == 11:
            self.Dialog14()

    def Dialog4(self):
        self.dialog4 = QDialog()
        self.dialog4.setStyleSheet('background-color: #2E2E38')
        self.dialog4.setWindowIcon(QIcon('./EY_logo.png'))

        # 트리 작업
        cursor = self.cnxn.cursor()

        sql = '''
                 SELECT 											
                        *
                 FROM  [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											

            '''.format(field=self.selected_project_id)

        accountsname = pd.read_sql(sql, self.cnxn)

        self.new_tree = Form(self)

        self.new_tree.tree.clear()

        for n, i in enumerate(accountsname.AccountType.unique()):
            self.new_tree.parent = QTreeWidgetItem(self.new_tree.tree)

            self.new_tree.parent.setText(0, "{}".format(i))
            self.new_tree.parent.setFlags(self.new_tree.parent.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
            child_items = accountsname.AccountSubType[
                accountsname.AccountType == accountsname.AccountType.unique()[n]].unique()
            for m, x in enumerate(child_items):
                self.new_tree.child = QTreeWidgetItem(self.new_tree.parent)

                self.new_tree.child.setText(0, "{}".format(x))
                self.new_tree.child.setFlags(self.new_tree.child.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                grandchild_items = accountsname.AccountClass[accountsname.AccountSubType == child_items[m]].unique()
                for o, y in enumerate(grandchild_items):
                    self.new_tree.grandchild = QTreeWidgetItem(self.new_tree.child)

                    self.new_tree.grandchild.setText(0, "{}".format(y))
                    self.new_tree.grandchild.setFlags(
                        self.new_tree.grandchild.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                    num_name = accountsname[accountsname.AccountClass == grandchild_items[o]].iloc[:, 2:4]
                    full_name = num_name["GLAccountNumber"].map(str) + ' ' + num_name["GLAccountName"]
                    for z in full_name:
                        self.new_tree.grandgrandchild = QTreeWidgetItem(self.new_tree.grandchild)

                        self.new_tree.grandgrandchild.setText(0, "{}".format(z))
                        self.new_tree.grandgrandchild.setFlags(
                            self.new_tree.grandgrandchild.flags() | Qt.ItemIsUserCheckable)
                        self.new_tree.grandgrandchild.setCheckState(0, Qt.Checked)
        self.new_tree.get_selected_leaves()  # 초기값 모두 선택 (추가)

        ### 버튼 1 - Extract Data
        self.btn2 = QPushButton('   Extract Data', self.dialog4)
        self.btn2.setStyleSheet('color:white; background-image : url(./bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked4)
        self.btn2.clicked.connect(self.doAction)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        ### 버튼 2 - Close
        self.btnDialog = QPushButton('   Close', self.dialog4)
        self.btnDialog.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close4)

        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)

        # JE Line Number / JE Number 라디오 버튼
        self.rbtn1 = QRadioButton('JE Line Number', self.dialog4)
        self.rbtn1.setStyleSheet("color: white;")
        font11 = self.rbtn1.font()
        font11.setBold(True)
        self.rbtn1.setFont(font11)
        self.rbtn1.setChecked(True)
        self.rbtn2 = QRadioButton('JE Number', self.dialog4)
        self.rbtn2.setStyleSheet("color: white;")
        font12 = self.rbtn2.font()
        font12.setBold(True)
        self.rbtn2.setFont(font12)

        ### Progress Bar
        self.progressBar = QProgressBar(self.dialog4)
        self.progressBar.setMaximum(100)
        self.progressBar.setMinimum(0)

        ### 라벨 1 - 사용빈도
        label_freq = QLabel('사용 빈도(N)* :', self.dialog4)
        label_freq.setStyleSheet('color: white;')

        font1 = label_freq.font()
        font1.setBold(True)
        label_freq.setFont(font1)

        ### LineEdit 1 - 사용 빈도
        self.D4_N = QLineEdit(self.dialog4)
        self.D4_N.setStyleSheet('background-color: white;')
        self.D4_N.setPlaceholderText('사용빈도를 입력하세요')

        ### 라벨 2 - 중요성 금액
        label_TE = QLabel('중요성 금액: ', self.dialog4)
        label_TE.setStyleSheet('color: white;')

        font2 = label_TE.font()
        font2.setBold(True)
        label_TE.setFont(font2)

        ### LineEdit 2 - 중요성 금액
        self.D4_TE = QLineEdit(self.dialog4)
        self.D4_TE.setStyleSheet('background-color: white;')
        self.D4_TE.setPlaceholderText('중요성 금액을 입력하세요')

        ### 라벨 3 - 시트명
        labelSheet = QLabel('시트명* : ', self.dialog4)
        labelSheet.setStyleSheet("color: white;")

        font5 = labelSheet.font()
        font5.setBold(True)
        labelSheet.setFont(font5)

        ### LineEdit 3 - 시트명
        self.D4_Sheet = QLineEdit(self.dialog4)
        self.D4_Sheet.setStyleSheet("background-color: white;")
        self.D4_Sheet.setPlaceholderText('시트명을 입력하세요')

        label_tree = QLabel('특정 계정명 : ', self.dialog4)
        label_tree.setStyleSheet("color: white;")
        font4 = label_tree.font()
        font4.setBold(True)
        label_tree.setFont(font4)

        self.D4_N.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D4_TE.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D4_Sheet.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        ### 요소 배치
        layout1 = QGridLayout()
        layout1.addWidget(self.rbtn1, 0, 0)
        layout1.addWidget(self.rbtn2, 0, 1)
        layout1.addWidget(label_freq, 1, 0)
        layout1.addWidget(self.D4_N, 1, 1)
        layout1.addWidget(label_TE, 2, 0)
        layout1.addWidget(self.D4_TE, 2, 1)
        layout1.addWidget(label_tree, 3, 0)
        layout1.addWidget(self.new_tree, 3, 1)
        layout1.addWidget(labelSheet, 4, 0)
        layout1.addWidget(self.D4_Sheet, 4, 1)

        layout2 = QHBoxLayout()
        layout2.addWidget(self.progressBar)
        layout2.addWidget(self.btn2)
        layout2.addWidget(self.btnDialog)

        layout2.setContentsMargins(-1, 10, -1, -1)

        main_layout = QVBoxLayout()
        main_layout.setAlignment(Qt.AlignTop)
        main_layout.addLayout(layout1)
        main_layout.addLayout(layout2)

        self.dialog4.setLayout(main_layout)
        self.dialog4.setGeometry(300, 300, 500, 150)

        # ? 제거
        self.dialog4.setWindowFlags(Qt.WindowCloseButtonHint)

        self.dialog4.setWindowTitle('Scenario4')
        self.dialog4.setWindowModality(Qt.NonModal)
        self.dialog4.show()

    def Dialog5(self):
        self.dialog5 = QDialog()
        self.dialog5.setStyleSheet('background-color: #2E2E38')
        self.dialog5.setWindowIcon(QIcon('./EY_logo.png'))

        cursor = self.cnxn.cursor()

        sql = '''
        SELECT *
        FROM  [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA
        '''.format(field=self.selected_project_id)

        accountsname = pd.read_sql(sql, self.cnxn)

        self.new_tree = Form(self)
        self.new_tree.tree.clear()

        for n, i in enumerate(accountsname.AccountType.unique()):
            self.new_tree.parent = QTreeWidgetItem(self.new_tree.tree)

            self.new_tree.parent.setText(0, "{}".format(i))
            self.new_tree.parent.setFlags(self.new_tree.parent.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
            child_items = accountsname.AccountSubType[
                accountsname.AccountType == accountsname.AccountType.unique()[n]].unique()
            for m, x in enumerate(child_items):
                self.new_tree.child = QTreeWidgetItem(self.new_tree.parent)

                self.new_tree.child.setText(0, "{}".format(x))
                self.new_tree.child.setFlags(self.new_tree.child.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                grandchild_items = accountsname.AccountClass[accountsname.AccountSubType == child_items[m]].unique()
                for o, y in enumerate(grandchild_items):
                    self.new_tree.grandchild = QTreeWidgetItem(self.new_tree.child)

                    self.new_tree.grandchild.setText(0, "{}".format(y))
                    self.new_tree.grandchild.setFlags(
                        self.new_tree.grandchild.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                    num_name = accountsname[accountsname.AccountClass == grandchild_items[o]].iloc[:, 2:4]
                    full_name = num_name["GLAccountNumber"].map(str) + ' ' + num_name["GLAccountName"]
                    for z in full_name:
                        self.new_tree.grandgrandchild = QTreeWidgetItem(self.new_tree.grandchild)

                        self.new_tree.grandgrandchild.setText(0, "{}".format(z))
                        self.new_tree.grandgrandchild.setFlags(
                            self.new_tree.grandgrandchild.flags() | Qt.ItemIsUserCheckable)
                        self.new_tree.grandgrandchild.setCheckState(0, Qt.Checked)
        self.new_tree.get_selected_leaves()  # 초기값 모두 선택 (추가)

        # JE Line Number / JE Number 라디오 버튼
        self.rbtn1 = QRadioButton('JE Line Number', self.dialog5)
        self.rbtn1.setStyleSheet("color: white;")
        font11 = self.rbtn1.font()
        font11.setBold(True)
        self.rbtn1.setFont(font11)
        self.rbtn1.setChecked(True)
        self.rbtn2 = QRadioButton('JE Number', self.dialog5)
        self.rbtn2.setStyleSheet("color: white;")
        font12 = self.rbtn2.font()
        font12.setBold(True)
        self.rbtn2.setFont(font12)

        ### 버튼 1 - Extract Data (Non-SAP)
        self.btn2 = QPushButton(' Extract Data', self.dialog5)
        self.btn2.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked5_Non_SAP)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        ### 버튼 2 - Extract Data (SAP)
        self.btn3 = QPushButton(' Extract Data', self.dialog5)
        self.btn3.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btn3.clicked.connect(self.extButtonClicked5_SAP)

        font11 = self.btn3.font()
        font11.setBold(True)
        self.btn3.setFont(font11)

        self.btn2.resize(110, 30)
        self.btn3.resize(110, 30)

        ### 버튼 3 - Close (SAP)
        self.btnDialog = QPushButton('Close', self.dialog5)
        self.btnDialog.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close5)

        font9 = self.btnDialog.font()
        font9.setBold(True)
        self.btnDialog.setFont(font9)

        ### 버튼 4 - Close (Non-SAP)
        self.btnDialog1 = QPushButton('Close', self.dialog5)
        self.btnDialog1.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btnDialog1.clicked.connect(self.dialog_close5)

        font11 = self.btnDialog1.font()
        font11.setBold(True)
        self.btnDialog1.setFont(font11)

        self.btnDialog.resize(110, 30)
        self.btnDialog1.resize(110, 30)

        ### SAP
        ### 라벨 1 - SKA1 파일 드롭하기
        label_SKA1 = QLabel('※ SKA1 파일을 Drop 하십시오', self.dialog5)
        label_SKA1.setStyleSheet('color: red;')

        font12 = label_SKA1.font()
        font12.setBold(False)
        label_SKA1.setFont(font12)

        ### 라벨 2 - YEAR (SAP)
        label_year = QLabel('연도* : ')
        label_year.setStyleSheet('color: white;')

        font13 = label_year.font()
        font13.setBold(True)
        label_year.setFont(font13)

        ### 라벨 3 - 시트명 (SAP)
        labelSheet = QLabel('시트명* : ', self.dialog5)
        labelSheet.setStyleSheet("color: white;")
        font5 = labelSheet.font()
        font5.setBold(True)
        labelSheet.setFont(font5)

        ### ListBox Widget 1 - SKA1
        self.listbox_drops = ListBoxWidget()
        self.listbox_drops.setStyleSheet('background-color: white;')

        ### LineEdit 1 - 시트명 (SAP)
        self.D5_Sheet = QLineEdit(self.dialog5)
        self.D5_Sheet.setStyleSheet("background-color: white;")
        self.D5_Sheet.setPlaceholderText('시트명을 입력하세요')

        ### LineEdit 2 - YEAR (SAP)
        self.D5_Year = QLineEdit(self.dialog5)
        self.D5_Year.setStyleSheet('background-color: white;')
        self.D5_Year.setPlaceholderText('연도를 입력하세요')

        ### Non-SAP
        ### 라벨1 - 계정코드 입력
        label_AccCode = QLabel('Enter your Account Code : ', self.dialog5)
        label_AccCode.setStyleSheet('color: white;')

        font1 = label_AccCode.font()
        font1.setBold(True)
        label_AccCode.setFont(font1)

        ### 라벨 2 - YEAR (Non SAP)
        label_year2 = QLabel('연도* : ')
        label_year2.setStyleSheet('color: white;')

        font14 = label_year2.font()
        font14.setBold(True)
        label_year2.setFont(font14)

        ### 라벨 3 - 시트명 (Non SAP)
        labelSheet2 = QLabel('시트명* : ', self.dialog5)
        labelSheet2.setStyleSheet("color: white;")
        font5 = labelSheet2.font()
        font5.setBold(True)
        labelSheet2.setFont(font5)

        ### TextEdit 1 - 계정코드 Paste
        self.MyInput = QTextEdit(self.dialog5)
        self.MyInput.setAcceptRichText(False)
        self.MyInput.setStyleSheet('background-color: white;')
        self.MyInput.setPlaceholderText('※ 입력 예시 : OO')

        ### LineEdit 1 - 시트명 (Non SAP)
        self.D5_Sheet2 = QLineEdit(self.dialog5)
        self.D5_Sheet2.setStyleSheet("background-color: white;")
        self.D5_Sheet2.setPlaceholderText('시트명을 입력하세요')

        ### LineEdit 2 - YEAR (Non SAP)
        self.D5_Year2 = QLineEdit(self.dialog5)
        self.D5_Year2.setStyleSheet('background-color: white;')
        self.D5_Year2.setPlaceholderText('연도를 입력하세요')

        ### Layout 구성
        layout = QVBoxLayout()

        layout1 = QVBoxLayout()
        sublayout1 = QVBoxLayout()
        sublayout2 = QHBoxLayout()
        sublayout5 = QGridLayout()

        layout2 = QVBoxLayout()
        sublayout3 = QVBoxLayout()
        sublayout4 = QHBoxLayout()
        sublayout6 = QGridLayout()

        layout3 = QVBoxLayout()
        sublayout7 = QVBoxLayout()
        sublayout8 = QHBoxLayout()
        sublayout9 = QGridLayout()

        layout0 = QGridLayout()
        layout0.addWidget(self.rbtn1, 0, 0)
        layout0.addWidget(self.rbtn2, 0, 1)

        ### 탭
        tabs = QTabWidget()
        tab1 = QWidget()
        tab2 = QWidget()

        tab1.setLayout(layout1)
        tab2.setLayout(layout2)

        tabs.addTab(tab1, "Non SAP")
        tabs.addTab(tab2, "SAP")

        layout.addLayout(layout0)
        layout.addWidget(tabs)

        ### 배치 - 탭 1
        sublayout1.addWidget(label_year)
        sublayout1.addWidget(self.D5_Year)
        sublayout1.addWidget(label_AccCode)
        sublayout1.addWidget(self.MyInput)

        sublayout5.addWidget(labelSheet, 2, 0)
        sublayout5.addWidget(self.D5_Sheet, 2, 1)

        layout1.addLayout(sublayout1, stretch=4)
        layout1.addLayout(sublayout5, stretch=4)
        layout1.addLayout(sublayout2, stretch=1)

        sublayout2.addStretch(2)
        sublayout2.addWidget(self.btn2)
        sublayout2.addWidget(self.btnDialog)

        ### 배치 - 탭 2
        sublayout3.addWidget(label_year2)
        sublayout3.addWidget(self.D5_Year2)
        sublayout3.addWidget(label_SKA1)
        sublayout3.addWidget(self.listbox_drops)

        sublayout6.addWidget(labelSheet2, 3, 0)
        sublayout6.addWidget(self.D5_Sheet2, 3, 1)

        layout2.addLayout(sublayout3, stretch=4)
        layout2.addLayout(sublayout6, stretch=4)
        layout2.addLayout(sublayout4, stretch=1)

        sublayout4.addStretch(2)
        sublayout4.addWidget(self.btn3)
        sublayout4.addWidget(self.btnDialog1)

        # ? 제거
        self.dialog5.setWindowFlags(Qt.WindowCloseButtonHint)

        ### 공통 지정
        self.dialog5.setLayout(layout)
        self.dialog5.resize(465, 400)
        self.dialog5.setWindowTitle('Scenario5')
        self.dialog5.setWindowModality(Qt.NonModal)
        self.dialog5.show()

    def Dialog6(self):
        self.dialog6 = QDialog()
        self.dialog6.setStyleSheet('background-color: #2E2E38')
        self.dialog6.setWindowIcon(QIcon("./EY_logo.png"))

        # Extraction 내 Dictionary 를 위한 변수 설정
        self.D6_clickcount = 0

        # CoA 트리 만들기
        cursor = self.cnxn.cursor()

        sql = '''
                 SELECT 											
                        *
                 FROM  [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											

            '''.format(field=self.selected_project_id)

        accountsname = pd.read_sql(sql, self.cnxn)
        self.new_tree = Form(self)
        self.new_tree.tree.clear()

        for n, i in enumerate(accountsname.AccountType.unique()):
            self.new_tree.parent = QTreeWidgetItem(self.new_tree.tree)

            self.new_tree.parent.setText(0, "{}".format(i))
            self.new_tree.parent.setFlags(self.new_tree.parent.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
            child_items = accountsname.AccountSubType[
                accountsname.AccountType == accountsname.AccountType.unique()[n]].unique()
            for m, x in enumerate(child_items):
                self.new_tree.child = QTreeWidgetItem(self.new_tree.parent)

                self.new_tree.child.setText(0, "{}".format(x))
                self.new_tree.child.setFlags(self.new_tree.child.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                grandchild_items = accountsname.AccountClass[accountsname.AccountSubType == child_items[m]].unique()
                for o, y in enumerate(grandchild_items):
                    self.new_tree.grandchild = QTreeWidgetItem(self.new_tree.child)

                    self.new_tree.grandchild.setText(0, "{}".format(y))
                    self.new_tree.grandchild.setFlags(
                        self.new_tree.grandchild.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                    num_name = accountsname[accountsname.AccountClass == grandchild_items[o]].iloc[:, 2:4]
                    full_name = num_name["GLAccountNumber"].map(str) + ' ' + num_name["GLAccountName"]
                    for z in full_name:
                        self.new_tree.grandgrandchild = QTreeWidgetItem(self.new_tree.grandchild)

                        self.new_tree.grandgrandchild.setText(0, "{}".format(z))
                        self.new_tree.grandgrandchild.setFlags(
                            self.new_tree.grandgrandchild.flags() | Qt.ItemIsUserCheckable)
                        self.new_tree.grandgrandchild.setCheckState(0, Qt.Checked)

        self.new_tree.get_selected_leaves()  # 초기값 모두 선택 (추가)

        # 데이터 추출 버튼
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

        # JE Line Number / JE Number 라디오 버튼
        self.rbtn1 = QRadioButton('JE Line Number', self.dialog6)
        self.rbtn1.setStyleSheet("color: white;")
        font11 = self.rbtn1.font()
        font11.setBold(True)
        self.rbtn1.setFont(font11)
        self.rbtn1.setChecked(True)
        self.rbtn2 = QRadioButton('JE Number', self.dialog6)
        self.rbtn2.setStyleSheet("color: white;")
        font12 = self.rbtn2.font()
        font12.setBold(True)
        self.rbtn2.setFont(font12)

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
        self.btnDate.clicked.connect(self.calendar6)
        font11 = self.btnDate.font()
        font11.setBold(True)
        self.btnDate.setFont(font11)

        labelDate2 = QLabel('T일 : ', self.dialog6)
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
        self.D6_Cost.setPlaceholderText('중요성 금액을 입력하세요')

        labelSheet = QLabel('시트명* : ', self.dialog6)
        labelSheet.setStyleSheet("color: white;")

        font5 = labelSheet.font()
        font5.setBold(True)
        labelSheet.setFont(font5)

        self.D6_Sheet = QLineEdit(self.dialog6)
        self.D6_Sheet.setStyleSheet("background-color: white;")
        self.D6_Sheet.setPlaceholderText('시트명을 입력하세요')

        self.D6_Date.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D6_Date2.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D6_JE.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D6_Cost.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D6_Sheet.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        layout1 = QGridLayout()
        layout1.addWidget(self.rbtn1, 0, 0)
        layout1.addWidget(self.rbtn2, 0, 1)
        layout1.addWidget(labelDate, 1, 0)
        layout1.addWidget(self.D6_Date, 1, 1)
        layout1.addWidget(self.btnDate, 1, 2)
        layout1.addWidget(labelDate2, 2, 0)
        layout1.addWidget(self.D6_Date2, 2, 1)
        layout1.addWidget(label_tree, 3, 0)
        layout1.addWidget(self.new_tree, 3, 1)
        layout1.addWidget(labelJE, 4, 0)
        layout1.addWidget(self.D6_JE, 4, 1)
        layout1.addWidget(labelCost, 5, 0)
        layout1.addWidget(self.D6_Cost, 5, 1)
        layout1.addWidget(labelSheet, 6, 0)
        layout1.addWidget(self.D6_Sheet, 6, 1)

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
        self.dialog6.setGeometry(300, 300, 700, 450)

        self.dialog6.setWindowFlags(Qt.WindowCloseButtonHint)

        self.dialog6.setWindowTitle("Scenario6")
        self.dialog6.setWindowModality(Qt.NonModal)
        self.dialog6.show()

    def Dialog7(self):
        self.dialog7 = QDialog()
        self.dialog7.setStyleSheet('background-color: #2E2E38')
        self.dialog7.setWindowIcon(QIcon("./EY_logo.png"))

        groupbox1 = QGroupBox('')
        groupbox1.setStyleSheet('QGroupBox  {border:0;}')
        font_groupbox1 = groupbox1.font()
        font_groupbox1.setBold(True)
        groupbox1.setFont(font_groupbox1)

        groupbox2 = QGroupBox('')
        groupbox2.setStyleSheet('QGroupBox  {border:0;}')
        font_groupbox2 = groupbox2.font()
        font_groupbox2.setBold(True)
        groupbox2.setFont(font_groupbox2)

        groupbox3 = QGroupBox('')
        groupbox3.setStyleSheet('QGroupBox  {border:0;}')
        font_groupbox3 = groupbox3.font()
        font_groupbox3.setBold(True)
        groupbox3.setFont(font_groupbox3)

        groupbox4 = QGroupBox('')
        groupbox4.setStyleSheet('QGroupBox  {border:0;}')
        font_groupbox4 = groupbox4.font()
        font_groupbox4.setBold(True)
        groupbox4.setFont(font_groupbox4)

        cursor = self.cnxn.cursor()

        sql = '''
                 SELECT 											
                        *
                 FROM  [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											

            '''.format(field=self.selected_project_id)

        accountsname = pd.read_sql(sql, self.cnxn)

        self.new_tree = Form(self)

        self.new_tree.tree.clear()

        for n, i in enumerate(accountsname.AccountType.unique()):
            self.new_tree.parent = QTreeWidgetItem(self.new_tree.tree)

            self.new_tree.parent.setText(0, "{}".format(i))
            self.new_tree.parent.setFlags(self.new_tree.parent.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
            child_items = accountsname.AccountSubType[
                accountsname.AccountType == accountsname.AccountType.unique()[n]].unique()
            for m, x in enumerate(child_items):
                self.new_tree.child = QTreeWidgetItem(self.new_tree.parent)

                self.new_tree.child.setText(0, "{}".format(x))
                self.new_tree.child.setFlags(self.new_tree.child.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                grandchild_items = accountsname.AccountClass[accountsname.AccountSubType == child_items[m]].unique()
                for o, y in enumerate(grandchild_items):
                    self.new_tree.grandchild = QTreeWidgetItem(self.new_tree.child)

                    self.new_tree.grandchild.setText(0, "{}".format(y))
                    self.new_tree.grandchild.setFlags(
                        self.new_tree.grandchild.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                    num_name = accountsname[accountsname.AccountClass == grandchild_items[o]].iloc[:, 2:4]
                    full_name = num_name["GLAccountNumber"].map(str) + ' ' + num_name["GLAccountName"]
                    for z in full_name:
                        self.new_tree.grandgrandchild = QTreeWidgetItem(self.new_tree.grandchild)

                        self.new_tree.grandgrandchild.setText(0, "{}".format(z))
                        self.new_tree.grandgrandchild.setFlags(
                            self.new_tree.grandgrandchild.flags() | Qt.ItemIsUserCheckable)
                        self.new_tree.grandgrandchild.setCheckState(0, Qt.Checked)
        self.new_tree.get_selected_leaves()  # 초기값 모두 선택 (추가)

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

        # JE Line Number / JE Number 라디오 버튼
        self.rbtn3 = QRadioButton('JE Line Number', self.dialog7)
        self.rbtn3.setStyleSheet("color: white;")
        font11 = self.rbtn3.font()
        font11.setBold(True)
        self.rbtn3.setFont(font11)

        self.rbtn4 = QRadioButton('JE Number', self.dialog7)
        self.rbtn4.setStyleSheet("color: white;")
        font12 = self.rbtn4.font()
        font12.setBold(True)
        self.rbtn4.setFont(font12)

        self.rbtn1 = QRadioButton('Effective Date', self.dialog7)
        self.rbtn1.setStyleSheet("color: white;")
        font1 = self.rbtn1.font()
        font1.setBold(True)
        self.rbtn1.setFont(font1)

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
        self.btnDate.clicked.connect(self.calendar7)

        font11 = self.btnDate.font()
        font11.setBold(True)
        self.btnDate.setFont(font11)

        label_tree = QLabel('특정 계정명 : ', self.dialog7)
        label_tree.setStyleSheet("color: white;")
        font4 = label_tree.font()
        font4.setBold(True)
        label_tree.setFont(font4)

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
        self.D7_Cost.setPlaceholderText('중요성 금액을 입력하세요')

        labelSheet = QLabel('시트명* : ', self.dialog7)
        labelSheet.setStyleSheet("color: white;")

        font5 = labelSheet.font()
        font5.setBold(True)
        labelSheet.setFont(font5)

        self.D7_Sheet = QLineEdit(self.dialog7)
        self.D7_Sheet.setStyleSheet("background-color: white;")
        self.D7_Sheet.setPlaceholderText('시트명을 입력하세요')

        self.D7_Date.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D7_JE.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D7_Cost.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D7_Sheet.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        layout0 = QGridLayout()
        self.rbtn1.setChecked(True)
        layout0.addWidget(self.rbtn1, 0, 0)
        layout0.addWidget(self.rbtn2, 0, 1)
        groupbox1.setLayout(layout0)

        layout1 = QGridLayout()
        layout1.addWidget(labelDate, 0, 0)
        layout1.addWidget(self.D7_Date, 0, 1)
        layout1.addWidget(self.btnDate, 0, 2)
        layout1.addWidget(label_tree, 1, 0)
        layout1.addWidget(self.new_tree, 1, 1)
        layout1.addWidget(labelJE, 2, 0)
        layout1.addWidget(self.D7_JE, 2, 1)
        layout1.addWidget(labelCost, 3, 0)
        layout1.addWidget(self.D7_Cost, 3, 1)
        layout1.addWidget(labelSheet, 4, 0)
        layout1.addWidget(self.D7_Sheet, 4, 1)
        groupbox3.setLayout(layout1)

        layout2 = QHBoxLayout()
        layout2.addStretch()
        layout2.addStretch()
        layout2.addWidget(self.btn2)
        layout2.addWidget(self.btnDialog)
        layout2.setContentsMargins(-1, 10, -1, -1)
        groupbox4.setLayout(layout2)

        layout3 = QGridLayout()
        self.rbtn3.setChecked(True)
        layout3.addWidget(self.rbtn3, 0, 0)
        layout3.addWidget(self.rbtn4, 0, 1)
        groupbox2.setLayout(layout3)

        main_layout = QVBoxLayout()
        main_layout.setAlignment(Qt.AlignTop)
        main_layout.addWidget(groupbox2)
        main_layout.addWidget(groupbox1)
        main_layout.addWidget(groupbox3)
        main_layout.addWidget(groupbox4)

        self.dialog7.setLayout(main_layout)
        self.dialog7.setGeometry(300, 300, 700, 400)

        self.dialog7.setWindowFlags(Qt.WindowCloseButtonHint)  # ? 제거

        self.dialog7.setWindowTitle("Scenario7")
        self.dialog7.setWindowModality(Qt.NonModal)
        self.dialog7.show()

    def Dialog8(self):
        self.dialog8 = QDialog()
        self.dialog8.setStyleSheet('background-color: #2E2E38')
        self.dialog8.setWindowIcon(QIcon("./EY_logo.png"))

        cursor = self.cnxn.cursor()

        sql = '''
                 SELECT 											
                        *
                 FROM  [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											

            '''.format(field=self.selected_project_id)

        accountsname = pd.read_sql(sql, self.cnxn)

        self.new_tree = Form(self)

        self.new_tree.tree.clear()

        for n, i in enumerate(accountsname.AccountType.unique()):
            self.new_tree.parent = QTreeWidgetItem(self.new_tree.tree)

            self.new_tree.parent.setText(0, "{}".format(i))
            self.new_tree.parent.setFlags(self.new_tree.parent.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
            child_items = accountsname.AccountSubType[
                accountsname.AccountType == accountsname.AccountType.unique()[n]].unique()
            for m, x in enumerate(child_items):
                self.new_tree.child = QTreeWidgetItem(self.new_tree.parent)

                self.new_tree.child.setText(0, "{}".format(x))
                self.new_tree.child.setFlags(self.new_tree.child.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                grandchild_items = accountsname.AccountClass[accountsname.AccountSubType == child_items[m]].unique()
                for o, y in enumerate(grandchild_items):
                    self.new_tree.grandchild = QTreeWidgetItem(self.new_tree.child)

                    self.new_tree.grandchild.setText(0, "{}".format(y))
                    self.new_tree.grandchild.setFlags(
                        self.new_tree.grandchild.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                    num_name = accountsname[accountsname.AccountClass == grandchild_items[o]].iloc[:, 2:4]
                    full_name = num_name["GLAccountNumber"].map(str) + ' ' + num_name["GLAccountName"]
                    for z in full_name:
                        self.new_tree.grandgrandchild = QTreeWidgetItem(self.new_tree.grandchild)

                        self.new_tree.grandgrandchild.setText(0, "{}".format(z))
                        self.new_tree.grandgrandchild.setFlags(
                            self.new_tree.grandgrandchild.flags() | Qt.ItemIsUserCheckable)
                        self.new_tree.grandgrandchild.setCheckState(0, Qt.Checked)
        self.new_tree.get_selected_leaves()  # 초기값 모두 선택 (추가)

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

        # JE Line Number / JE Number 라디오 버튼
        self.rbtn1 = QRadioButton('JE Line Number', self.dialog8)
        self.rbtn1.setStyleSheet("color: white;")
        font11 = self.rbtn1.font()
        font11.setBold(True)
        self.rbtn1.setFont(font11)
        self.rbtn1.setChecked(True)
        self.rbtn2 = QRadioButton('JE Number', self.dialog8)
        self.rbtn2.setStyleSheet("color: white;")
        font12 = self.rbtn2.font()
        font12.setBold(True)
        self.rbtn2.setFont(font12)

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
        self.D8_Cost.setPlaceholderText('중요성 금액을 입력하세요')

        labelSheet = QLabel('시트명* : ', self.dialog8)
        labelSheet.setStyleSheet("color: white;")

        font5 = labelSheet.font()
        font5.setBold(True)
        labelSheet.setFont(font5)

        self.D8_Sheet = QLineEdit(self.dialog8)
        self.D8_Sheet.setStyleSheet("background-color: white;")
        self.D8_Sheet.setPlaceholderText('시트명을 입력하세요')

        self.D8_N.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D8_JE.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D8_Cost.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D8_Sheet.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        layout1 = QGridLayout()
        layout1.addWidget(self.rbtn1, 0, 0)
        layout1.addWidget(self.rbtn2, 0, 1)
        layout1.addWidget(labelDate, 1, 0)
        layout1.addWidget(self.D8_N, 1, 1)
        layout1.addWidget(label_tree, 2, 0)
        layout1.addWidget(self.new_tree, 2, 1)
        layout1.addWidget(labelJE, 3, 0)
        layout1.addWidget(self.D8_JE, 3, 1)
        layout1.addWidget(labelCost, 4, 0)
        layout1.addWidget(self.D8_Cost, 4, 1)
        layout1.addWidget(labelSheet, 5, 0)
        layout1.addWidget(self.D8_Sheet, 5, 1)

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

        # ? 제거
        self.dialog8.setWindowFlags(Qt.WindowCloseButtonHint)

        self.dialog8.setWindowTitle("Scenario8")
        self.dialog8.setWindowModality(Qt.NonModal)
        self.dialog8.show()

    def Dialog9(self):
        self.dialog9 = QDialog()
        groupbox = QGroupBox('접속 정보')
        cursor = self.cnxn.cursor()

        sql = '''
                 SELECT 											
                        *
                 FROM  [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											

            '''.format(field=self.selected_project_id)

        accountsname = pd.read_sql(sql, self.cnxn)

        self.new_tree = Form(self)

        self.new_tree.tree.clear()

        for n, i in enumerate(accountsname.AccountType.unique()):
            self.new_tree.parent = QTreeWidgetItem(self.new_tree.tree)

            self.new_tree.parent.setText(0, "{}".format(i))
            self.new_tree.parent.setFlags(self.new_tree.parent.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
            child_items = accountsname.AccountSubType[
                accountsname.AccountType == accountsname.AccountType.unique()[n]].unique()
            for m, x in enumerate(child_items):
                self.new_tree.child = QTreeWidgetItem(self.new_tree.parent)

                self.new_tree.child.setText(0, "{}".format(x))
                self.new_tree.child.setFlags(self.new_tree.child.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                grandchild_items = accountsname.AccountClass[accountsname.AccountSubType == child_items[m]].unique()
                for o, y in enumerate(grandchild_items):
                    self.new_tree.grandchild = QTreeWidgetItem(self.new_tree.child)

                    self.new_tree.grandchild.setText(0, "{}".format(y))
                    self.new_tree.grandchild.setFlags(
                        self.new_tree.grandchild.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                    num_name = accountsname[accountsname.AccountClass == grandchild_items[o]].iloc[:, 2:4]
                    full_name = num_name["GLAccountNumber"].map(str) + ' ' + num_name["GLAccountName"]
                    for z in full_name:
                        self.new_tree.grandgrandchild = QTreeWidgetItem(self.new_tree.grandchild)

                        self.new_tree.grandgrandchild.setText(0, "{}".format(z))
                        self.new_tree.grandgrandchild.setFlags(
                            self.new_tree.grandgrandchild.flags() | Qt.ItemIsUserCheckable)
                        self.new_tree.grandgrandchild.setCheckState(0, Qt.Checked)
        self.new_tree.get_selected_leaves()  # 초기값 모두 선택 (추가)

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

        # JE Line Number / JE Number 라디오 버튼
        self.rbtn1 = QRadioButton('JE Line Number', self.dialog9)
        self.rbtn1.setStyleSheet("color: white;")
        font11 = self.rbtn1.font()
        font11.setBold(True)
        self.rbtn1.setFont(font11)
        self.rbtn1.setChecked(True)
        self.rbtn2 = QRadioButton('JE Number', self.dialog9)
        self.rbtn2.setStyleSheet("color: white;")
        font12 = self.rbtn2.font()
        font12.setBold(True)
        self.rbtn2.setFont(font12)

        labelKeyword = QLabel('작성빈도수* : ', self.dialog9)
        labelKeyword.setStyleSheet("color: white;")

        font1 = labelKeyword.font()
        font1.setBold(True)
        labelKeyword.setFont(font1)

        self.D9_N = QLineEdit(self.dialog9)
        self.D9_N.setStyleSheet("background-color: white;")
        self.D9_N.setPlaceholderText('작성 빈도수를 입력하세요')

        labelTE = QLabel('TE : ', self.dialog9)
        labelTE.setStyleSheet("color: white;")

        font2 = labelTE.font()
        font2.setBold(True)
        labelTE.setFont(font2)

        self.D9_TE = QLineEdit(self.dialog9)
        self.D9_TE.setStyleSheet("background-color: white;")
        self.D9_TE.setPlaceholderText('중요성 금액을 입력하세요')

        self.btn2.resize(110, 30)
        self.btnDialog.resize(110, 30)

        label_tree = QLabel('특정 계정명 : ', self.dialog9)
        label_tree.setStyleSheet("color: white;")
        font4 = label_tree.font()
        font4.setBold(True)
        label_tree.setFont(font4)

        labelSheet = QLabel('시트명* : ', self.dialog9)
        labelSheet.setStyleSheet("color: white;")

        font5 = labelSheet.font()
        font5.setBold(True)
        labelSheet.setFont(font5)

        self.D9_Sheet = QLineEdit(self.dialog9)
        self.D9_Sheet.setStyleSheet("background-color: white;")
        self.D9_Sheet.setPlaceholderText('시트명을 입력하세요')

        self.D9_N.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D9_TE.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D9_Sheet.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        layout1 = QGridLayout()
        layout1.addWidget(self.rbtn1, 0, 0)
        layout1.addWidget(self.rbtn2, 0, 1)
        layout1.addWidget(labelKeyword, 1, 0)
        layout1.addWidget(self.D9_N, 1, 1)
        layout1.addWidget(labelTE, 2, 0)
        layout1.addWidget(self.D9_TE, 2, 1)
        layout1.addWidget(label_tree, 3, 0)
        layout1.addWidget(self.new_tree, 3, 1)
        layout1.addWidget(labelSheet, 4, 0)
        layout1.addWidget(self.D9_Sheet, 4, 1)

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

        self.dialog9.setLayout(main_layout)
        self.dialog9.setGeometry(300, 300, 500, 150)

        # ? 제거
        self.dialog9.setWindowFlags(Qt.WindowCloseButtonHint)

        self.dialog9.setWindowTitle("Scenario9")
        self.dialog9.setWindowModality(Qt.NonModal)
        self.dialog9.show()

    def Dialog10(self):
        self.dialog10 = QDialog()
        self.dialog10.setStyleSheet('background-color: #2E2E38')
        self.dialog10.setWindowIcon(QIcon("./EY_logo.png"))

        cursor = self.cnxn.cursor()

        sql = '''
                 SELECT 											
                        *
                 FROM  [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											

            '''.format(field=self.selected_project_id)

        accountsname = pd.read_sql(sql, self.cnxn)

        self.new_tree = Form(self)

        self.new_tree.tree.clear()

        for n, i in enumerate(accountsname.AccountType.unique()):
            self.new_tree.parent = QTreeWidgetItem(self.new_tree.tree)

            self.new_tree.parent.setText(0, "{}".format(i))
            self.new_tree.parent.setFlags(self.new_tree.parent.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
            child_items = accountsname.AccountSubType[
                accountsname.AccountType == accountsname.AccountType.unique()[n]].unique()
            for m, x in enumerate(child_items):
                self.new_tree.child = QTreeWidgetItem(self.new_tree.parent)

                self.new_tree.child.setText(0, "{}".format(x))
                self.new_tree.child.setFlags(self.new_tree.child.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                grandchild_items = accountsname.AccountClass[accountsname.AccountSubType == child_items[m]].unique()
                for o, y in enumerate(grandchild_items):
                    self.new_tree.grandchild = QTreeWidgetItem(self.new_tree.child)

                    self.new_tree.grandchild.setText(0, "{}".format(y))
                    self.new_tree.grandchild.setFlags(
                        self.new_tree.grandchild.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                    num_name = accountsname[accountsname.AccountClass == grandchild_items[o]].iloc[:, 2:4]
                    full_name = num_name["GLAccountNumber"].map(str) + ' ' + num_name["GLAccountName"]
                    for z in full_name:
                        self.new_tree.grandgrandchild = QTreeWidgetItem(self.new_tree.grandchild)

                        self.new_tree.grandgrandchild.setText(0, "{}".format(z))
                        self.new_tree.grandgrandchild.setFlags(
                            self.new_tree.grandgrandchild.flags() | Qt.ItemIsUserCheckable)
                        self.new_tree.grandgrandchild.setCheckState(0, Qt.Checked)
        self.new_tree.get_selected_leaves()  # 초기값 모두 선택 (추가)

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

        # JE Line Number / JE Number 라디오 버튼
        self.rbtn1 = QRadioButton('JE Line Number', self.dialog10)
        self.rbtn1.setStyleSheet("color: white;")
        font11 = self.rbtn1.font()
        font11.setBold(True)
        self.rbtn1.setFont(font11)
        self.rbtn1.setChecked(True)
        self.rbtn2 = QRadioButton('JE Number', self.dialog10)
        self.rbtn2.setStyleSheet("color: white;")
        font12 = self.rbtn2.font()
        font12.setBold(True)
        self.rbtn2.setFont(font12)

        labelKeyword = QLabel('전표입력자* : ', self.dialog10)
        labelKeyword.setStyleSheet("color: white;")

        font1 = labelKeyword.font()
        font1.setBold(True)
        labelKeyword.setFont(font1)

        self.D10_Search = QLineEdit(self.dialog10)
        self.D10_Search.setStyleSheet("background-color: white;")
        self.D10_Search.setPlaceholderText('전표입력자를 입력하세요')

        labelPoint1 = QLabel('시작시점 : ', self.dialog10)
        labelPoint2 = QLabel('종료시점 : ', self.dialog10)
        labelPoint1.setStyleSheet("color: white;")
        labelPoint2.setStyleSheet("color: white;")

        font2_1 = labelPoint1.font()
        font2_1.setBold(True)
        labelPoint1.setFont(font2_1)

        font2_2 = labelPoint2.font()
        font2_2.setBold(True)
        labelPoint2.setFont(font2_2)

        self.D10_Point1 = QLineEdit(self.dialog10)
        self.D10_Point1.setStyleSheet("background-color: white;")
        self.D10_Point1.setPlaceholderText('시작시점을 입력하세요')

        self.D10_Point2 = QLineEdit(self.dialog10)
        self.D10_Point2.setStyleSheet("background-color: white;")
        self.D10_Point2.setPlaceholderText('종료시점을 입력하세요')

        label_tree = QLabel('특정 계정명 : ', self.dialog10)
        label_tree.setStyleSheet("color: white;")
        font4 = label_tree.font()
        font4.setBold(True)
        label_tree.setFont(font4)

        labelTE = QLabel('TE : ', self.dialog10)
        labelTE.setStyleSheet("color: white;")

        font4 = labelTE.font()
        font4.setBold(True)
        labelTE.setFont(font4)

        self.D10_TE = QLineEdit(self.dialog10)
        self.D10_TE.setStyleSheet("background-color: white;")
        self.D10_TE.setPlaceholderText('중요성 금액을 입력하세요')

        labelSheet = QLabel('시트명* : ', self.dialog10)
        labelSheet.setStyleSheet("color: white;")

        font5 = labelSheet.font()
        font5.setBold(True)
        labelSheet.setFont(font5)

        self.D10_Sheet = QLineEdit(self.dialog10)
        self.D10_Sheet.setStyleSheet("background-color: white;")
        self.D10_Sheet.setPlaceholderText('시트명을 입력하세요')

        self.D10_Search.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D10_Point1.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D10_Point2.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D10_TE.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D10_Sheet.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        layout1 = QGridLayout()
        layout1.addWidget(self.rbtn1, 0, 0)
        layout1.addWidget(self.rbtn2, 0, 1)
        layout1.addWidget(labelKeyword, 1, 0)
        layout1.addWidget(self.D10_Search, 1, 1)
        layout1.addWidget(labelPoint1, 2, 0)
        layout1.addWidget(self.D10_Point1, 2, 1)
        layout1.addWidget(labelPoint2, 3, 0)
        layout1.addWidget(self.D10_Point2, 3, 1)
        layout1.addWidget(label_tree, 4, 0)
        layout1.addWidget(self.new_tree, 4, 1)
        layout1.addWidget(labelTE, 5, 0)
        layout1.addWidget(self.D10_TE, 5, 1)
        layout1.addWidget(labelSheet, 6, 0)
        layout1.addWidget(self.D10_Sheet, 6, 1)

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

        # ? 제거
        self.dialog10.setWindowFlags(Qt.WindowCloseButtonHint)

        self.dialog10.setWindowTitle("Scenario10")
        self.dialog10.setWindowModality(Qt.NonModal)
        self.dialog10.show()

    def Dialog11(self):
        self.dialog11 = QDialog()
        self.dialog11.setStyleSheet('background-color: #2E2E38')
        self.dialog11.setWindowIcon(QIcon('./EY_logo.png'))

        ### 버튼 1 - Extract Data Tab 1
        self.btn2 = QPushButton(' Extract Data', self.dialog11)
        self.btn2.setStyleSheet('color: white; background-image: url(./bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked11)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        ### 버튼 2 - Close Tab 1
        self.btnDialog = QPushButton('Close', self.dialog11)
        self.btnDialog.setStyleSheet('color: white; background-image: url(./bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close11)

        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)

        ### 버튼 3 - Extract Data Tab 2
        self.btn3 = QPushButton(' Extract Data', self.dialog11)
        self.btn3.setStyleSheet('color: white; background-image: url(./bar.png)')
        self.btn3.clicked.connect(self.extButtonClicked11)

        font11 = self.btn3.font()
        font11.setBold(True)
        self.btn3.setFont(font11)

        ### 버튼 4 - Close Tab 2
        self.btnDialog1 = QPushButton('Close', self.dialog11)
        self.btnDialog1.setStyleSheet('color: white; background-image: url(./bar.png)')
        self.btnDialog1.clicked.connect(self.dialog_close11)

        font13 = self.btnDialog1.font()
        font13.setBold(True)
        self.btnDialog1.setFont(font13)

        ### 라벨 1 - 계정명 (A)
        labelAccount_A = QLabel('주계정(A)을 선택하세요: ', self.dialog11)
        labelAccount_A.setStyleSheet('color: white;')

        font7 = labelAccount_A.font()
        font7.setBold(True)
        labelAccount_A.setFont(font7)

        ### 라벨 2 - 계정명 (B)
        labelAccount_B = QLabel('상대계정(B)을 선택하세요: ', self.dialog11)
        labelAccount_B.setStyleSheet('color: white;')

        font7 = labelAccount_B.font()
        font7.setBold(True)
        labelAccount_B.setFont(font7)

        ### 라벨 3 - 계정코드 (A)
        labelCode_A = QLabel('*주계정(A) 코드를 입력하세요:', self.dialog11)
        labelCode_A.setStyleSheet('color: red;')

        font11 = labelCode_A.font()
        font11.setBold(True)
        labelCode_A.setFont(font11)

        ### 라벨 4 - 계정코드 (B)
        labelCode_B = QLabel('*상대계정(B) 코드를 입력하세요:', self.dialog11)
        labelCode_B.setStyleSheet('color: red;')

        font11 = labelCode_B.font()
        font11.setBold(True)
        labelCode_B.setFont(font11)

        ### 라벨 5 - 중요성 금액 (A)
        labelTE_A = QLabel('중요성 금액 : ', self.dialog11)
        labelTE_A.setStyleSheet('color: white;')

        font7 = labelTE_A.font()
        font7.setBold(True)
        labelTE_A.setFont(font7)

        ### 라벨 6 - 중요성 금액 (B)
        labelTE_B = QLabel('중요성 금액 : ', self.dialog11)
        labelTE_B.setStyleSheet('color: white;')

        font11 = labelTE_B.font()
        font11.setBold(True)
        labelTE_B.setFont(font11)

        ### Line Edit 1 - 중요성 금액 (A)
        self.line_TE_A = QLineEdit(self.dialog11)
        self.line_TE_A.setStyleSheet('background-color: white;')
        self.line_TE_A.setPlaceholderText('중요성 금액을 입력하세요')

        ### Line Edit 2 - 중요성 금액 (B)
        self.line_TE_B = QLineEdit(self.dialog11)
        self.line_TE_B.setStyleSheet('background-color: white;')
        self.line_TE_B.setPlaceholderText('중요성 금액을 입력하세요')

        ### TextEdit - 계정코드 Paste (A)
        self.MyInput_A = QTextEdit(self.dialog11)
        self.MyInput_A.setAcceptRichText(False)
        self.MyInput_A.setStyleSheet('background-color: white;')
        self.MyInput_A.setPlaceholderText('계정코드를 입력하세요')

        ### TextEdit - 계정코드 Paste (B)
        self.MyInput_B = QTextEdit(self.dialog11)
        self.MyInput_B.setAcceptRichText(False)
        self.MyInput_B.setStyleSheet('background-color: white;')
        self.MyInput_B.setPlaceholderText('계정코드를 입력하세요')

        ### Layout - 다이얼로그 UI
        layout = QVBoxLayout()

        layout1 = QVBoxLayout()
        sublayout1 = QVBoxLayout()
        sublayout2 = QHBoxLayout()

        layout2 = QVBoxLayout()
        sublayout3 = QVBoxLayout()
        sublayout4 = QHBoxLayout()

        ### 탭
        my_tabs = QTabWidget()
        tab1 = QWidget()
        tab1.setLayout(layout1)
        my_tabs.addTab(tab1, "Account Name")

        tab2 = QWidget()
        tab2.setLayout(layout2)
        my_tabs.addTab(tab2, "Account Code")

        layout.addWidget(my_tabs)

        ### 요소 채우기 - 탭1
        sublayout1.addWidget(labelAccount_A)
        sublayout1.addWidget(self.account_tree_A)
        sublayout1.addWidget(labelAccount_B)
        sublayout1.addWidget(self.account_tree_B)
        sublayout1.addWidget(labelTE_A)
        sublayout1.addWidget(self.line_TE_A)

        sublayout2.addStretch(2)
        sublayout2.addWidget(self.btn2, stretch=1, alignment=Qt.AlignBottom)
        sublayout2.addWidget(self.btnDialog, stretch=1, alignment=Qt.AlignBottom)

        layout1.addLayout(sublayout1, stretch=4)
        layout1.addLayout(sublayout2, stretch=1)

        ### 요소 채우기 - 탭2
        sublayout3.addWidget(labelCode_A)
        sublayout3.addWidget(self.MyInput_A)
        sublayout3.addWidget(labelCode_B)
        sublayout3.addWidget(self.MyInput_B)
        sublayout3.addWidget(labelTE_B)
        sublayout3.addWidget(self.line_TE_B)

        sublayout4.addStretch(2)
        sublayout4.addWidget(self.btn3, stretch=1, alignment=Qt.AlignBottom)
        sublayout4.addWidget(self.btnDialog1, stretch=1, alignment=Qt.AlignBottom)

        layout2.addLayout(sublayout3, stretch=4)
        layout2.addLayout(sublayout4, stretch=1)

        self.dialog11.setLayout(layout)
        self.dialog11.resize(625, 550)

        # ? 제거
        self.dialog11.setWindowFlags(Qt.WindowCloseButtonHint)

        self.dialog11.setWindowTitle('Scenario 11')
        self.dialog11.setWindowModality(Qt.NonModal)
        self.dialog11.show()

    def Dialog12(self):  # 수정완료
        self.dialog12 = QDialog()
        self.dialog12.setStyleSheet('background-color: #2E2E38')
        self.dialog12.setWindowIcon(QIcon('./EY_logo.png'))

        cursor = self.cnxn.cursor()

        sql = '''
                         SELECT 											
                                *
                         FROM  [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											

                    '''.format(field=self.selected_project_id)

        accountsname = pd.read_sql(sql, self.cnxn)

        self.new_tree = Form(self)

        self.new_tree.tree.clear()

        for n, i in enumerate(accountsname.AccountType.unique()):
            self.new_tree.parent = QTreeWidgetItem(self.new_tree.tree)

            self.new_tree.parent.setText(0, "{}".format(i))
            self.new_tree.parent.setFlags(self.new_tree.parent.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
            child_items = accountsname.AccountSubType[
                accountsname.AccountType == accountsname.AccountType.unique()[n]].unique()
            for m, x in enumerate(child_items):
                self.new_tree.child = QTreeWidgetItem(self.new_tree.parent)

                self.new_tree.child.setText(0, "{}".format(x))
                self.new_tree.child.setFlags(self.new_tree.child.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                grandchild_items = accountsname.AccountClass[accountsname.AccountSubType == child_items[m]].unique()
                for o, y in enumerate(grandchild_items):
                    self.new_tree.grandchild = QTreeWidgetItem(self.new_tree.child)

                    self.new_tree.grandchild.setText(0, "{}".format(y))
                    self.new_tree.grandchild.setFlags(
                        self.new_tree.grandchild.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                    num_name = accountsname[accountsname.AccountClass == grandchild_items[o]].iloc[:, 2:4]
                    full_name = num_name["GLAccountNumber"].map(str) + ' ' + num_name["GLAccountName"]
                    for z in full_name:
                        self.new_tree.grandgrandchild = QTreeWidgetItem(self.new_tree.grandchild)

                        self.new_tree.grandgrandchild.setText(0, "{}".format(z))
                        self.new_tree.grandgrandchild.setFlags(
                            self.new_tree.grandgrandchild.flags() | Qt.ItemIsUserCheckable)
                        self.new_tree.grandgrandchild.setCheckState(0, Qt.Checked)
        self.new_tree.get_selected_leaves()

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

        # JE Line Number / JE Number 라디오 버튼
        self.rbtn1 = QRadioButton('JE Line Number', self.dialog12)
        self.rbtn1.setStyleSheet("color: white;")
        font11 = self.rbtn1.font()
        font11.setBold(True)
        self.rbtn1.setFont(font11)
        self.rbtn1.setChecked(True)
        self.rbtn2 = QRadioButton('JE Number', self.dialog12)
        self.rbtn2.setStyleSheet("color: white;")
        font12 = self.rbtn2.font()
        font12.setBold(True)
        self.rbtn2.setFont(font12)

        labelAccount = QLabel('특정 계정명/계정 코드* : ', self.dialog12)
        labelAccount.setStyleSheet("color: white;")
        font3 = labelAccount.font()
        font3.setBold(True)
        labelAccount.setFont(font3)

        labelCost = QLabel('중요성 금액 : ', self.dialog12)
        labelCost.setStyleSheet("color: white;")
        font3 = labelCost.font()
        font3.setBold(True)
        labelCost.setFont(font3)

        self.D12_Cost = QLineEdit(self.dialog12)
        self.D12_Cost.setStyleSheet("background-color: white;")
        self.D12_Cost.setPlaceholderText('100,000,000원 이상 입력하세요')
        self.D12_Cost.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        sublayout1 = QGridLayout()
        sublayout1.addWidget(self.rbtn1, 0, 0)
        sublayout1.addWidget(self.rbtn2, 0, 1)
        sublayout1.addWidget(labelAccount, 1, 0)
        sublayout1.addWidget(self.new_tree, 1, 1)
        sublayout1.addWidget(labelCost, 2, 0)
        sublayout1.addWidget(self.D12_Cost, 2, 1)

        sublayout2 = QHBoxLayout()
        sublayout2.addStretch()
        sublayout2.addStretch()
        sublayout2.addWidget(self.btn)
        sublayout2.addWidget(self.btnDialog)

        main_layout = QVBoxLayout()
        main_layout.addLayout(sublayout1)
        main_layout.addStretch()
        main_layout.addLayout(sublayout2)

        self.dialog12.setLayout(main_layout)
        self.dialog12.resize(500, 300)

        # ? 제거
        self.dialog12.setWindowFlags(Qt.WindowCloseButtonHint)

        self.dialog12.setWindowTitle('Scenario12')
        self.dialog12.setWindowModality(Qt.NonModal)
        self.dialog12.show()

    def Dialog13(self):
        self.dialog13 = QDialog()
        self.dialog13.setStyleSheet('background-color: #2E2E38')
        self.dialog13.setWindowIcon(QIcon('./EY_logo.png'))

        cursor = self.cnxn.cursor()

        sql = '''
                 SELECT 											
                        *
                 FROM  [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											

            '''.format(field=self.selected_project_id)

        accountsname = pd.read_sql(sql, self.cnxn)

        self.new_tree = Form(self)

        self.new_tree.tree.clear()

        for n, i in enumerate(accountsname.AccountType.unique()):
            self.new_tree.parent = QTreeWidgetItem(self.new_tree.tree)

            self.new_tree.parent.setText(0, "{}".format(i))
            self.new_tree.parent.setFlags(self.new_tree.parent.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
            child_items = accountsname.AccountSubType[
                accountsname.AccountType == accountsname.AccountType.unique()[n]].unique()
            for m, x in enumerate(child_items):
                self.new_tree.child = QTreeWidgetItem(self.new_tree.parent)

                self.new_tree.child.setText(0, "{}".format(x))
                self.new_tree.child.setFlags(self.new_tree.child.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                grandchild_items = accountsname.AccountClass[accountsname.AccountSubType == child_items[m]].unique()
                for o, y in enumerate(grandchild_items):
                    self.new_tree.grandchild = QTreeWidgetItem(self.new_tree.child)

                    self.new_tree.grandchild.setText(0, "{}".format(y))
                    self.new_tree.grandchild.setFlags(
                        self.new_tree.grandchild.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                    num_name = accountsname[accountsname.AccountClass == grandchild_items[o]].iloc[:, 2:4]
                    full_name = num_name["GLAccountNumber"].map(str) + ' ' + num_name["GLAccountName"]
                    for z in full_name:
                        self.new_tree.grandgrandchild = QTreeWidgetItem(self.new_tree.grandchild)

                        self.new_tree.grandgrandchild.setText(0, "{}".format(z))
                        self.new_tree.grandgrandchild.setFlags(
                            self.new_tree.grandgrandchild.flags() | Qt.ItemIsUserCheckable)
                        self.new_tree.grandgrandchild.setCheckState(0, Qt.Checked)
        self.new_tree.get_selected_leaves()  # 초기값 모두 선택 (추가)

        # JE Line Number / JE Number 라디오 버튼
        self.rbtn1 = QRadioButton('JE Line Number', self.dialog13)
        self.rbtn1.setStyleSheet("color: white;")
        font11 = self.rbtn1.font()
        font11.setBold(True)
        self.rbtn1.setFont(font11)
        self.rbtn1.setChecked(True)
        self.rbtn2 = QRadioButton('JE Number', self.dialog13)
        self.rbtn2.setStyleSheet("color: white;")
        font12 = self.rbtn2.font()
        font12.setBold(True)
        self.rbtn2.setFont(font12)

        ### 버튼 1 - Extract Data
        self.btn2 = QPushButton(' Extract Data', self.dialog13)
        self.btn2.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked13)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        ### 버튼 2 - Close
        self.btnDialog = QPushButton('Close', self.dialog13)
        self.btnDialog.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close13)

        font9 = self.btnDialog.font()
        font9.setBold(True)
        self.btnDialog.setFont(font9)

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

        font13 = self.btnSaveProceed.font()
        font13.setBold(True)
        self.btnSaveProceed.setFont(font13)

        ### 라벨 1 - 연속된 자릿수
        label_Continuous = QLabel('* 연속된 자릿수 (ex. 3333, 6666): ', self.dialog13)
        label_Continuous.setStyleSheet("color: red;")

        font1 = label_Continuous.font()
        font1.setBold(True)
        label_Continuous.setFont(font1)

        ### Text Edit - 연속된 자릿수
        self.text_continuous = QTextEdit(self.dialog13)
        self.text_continuous.setAcceptRichText(False)
        self.text_continuous.setStyleSheet("background-color: white;")

        ### 라벨 2 - 중요성 금액
        label_amount = QLabel('중요성 금액 : ', self.dialog13)
        label_amount.setStyleSheet("color: white;")

        font3 = label_amount.font()
        font3.setBold(True)
        label_amount.setFont(font3)

        ### Line Edit - 중요성 금액
        self.line_amount = QLineEdit(self.dialog13)
        self.line_amount.setStyleSheet("background-color: white;")

        ### 라벨 3 - 계정 트리
        label_tree = QLabel('원하는 계정명을 선택하세요 : ', self.dialog13)
        label_tree.setStyleSheet("color: white;")
        font4 = label_tree.font()
        font4.setBold(True)
        label_tree.setFont(font4)

        labelSheet = QLabel('시트명* : ', self.dialog13)
        labelSheet.setStyleSheet("color: white;")

        font5 = labelSheet.font()
        font5.setBold(True)
        labelSheet.setFont(font5)

        self.D13_Sheet = QLineEdit(self.dialog13)
        self.D13_Sheet.setStyleSheet("background-color: white;")
        self.D13_Sheet.setPlaceholderText('시트명을 입력하세요')

        ### Layout - 다이얼로그 UI
        main_layout = QVBoxLayout()

        layout1 = QVBoxLayout()
        sublayout1 = QVBoxLayout()
        sublayout2 = QHBoxLayout()

        layout2 = QVBoxLayout()
        sublayout3 = QVBoxLayout()
        sublayout4 = QHBoxLayout()
        sublayout5 = QGridLayout()

        layout0 = QGridLayout()
        layout0.addWidget(self.rbtn1, 0, 0)
        layout0.addWidget(self.rbtn2, 0, 1)

        ### 탭
        tabs = QTabWidget()
        tab1 = QWidget()
        tab2 = QWidget()

        tab1.setLayout(layout1)
        tab2.setLayout(layout2)
        tabs.addTab(tab1, "Step 1")
        tabs.addTab(tab2, "Step 2")

        main_layout.addLayout(layout0)
        main_layout.addWidget(tabs)

        ### 배치 - 탭 1
        sublayout1.addWidget(label_Continuous)
        sublayout1.addWidget(self.text_continuous)
        sublayout1.addWidget(label_amount)
        sublayout1.addWidget(self.line_amount)

        sublayout2.addStretch(2)
        sublayout2.addWidget(self.btnSave)
        sublayout2.addWidget(self.btnSaveProceed)

        layout1.addLayout(sublayout1, stretch=4)
        layout1.addLayout(sublayout2, stretch=1)

        ### 배치 - 탭 2
        sublayout3.addWidget(label_tree)
        sublayout3.addWidget(self.new_tree)

        sublayout5.addWidget(labelSheet, 0, 0)
        sublayout5.addWidget(self.D13_Sheet, 0, 1)

        sublayout4.addStretch(2)
        sublayout4.addWidget(self.btn2)
        sublayout4.addWidget(self.btnDialog)

        layout2.addLayout(sublayout3, stretch=4)
        layout2.addLayout(sublayout5, stretch=1)
        layout2.addLayout(sublayout4, stretch=1)

        ### 공통 지정
        self.dialog13.setLayout(main_layout)
        self.dialog13.resize(500, 400)

        # ? 제거
        self.dialog13.setWindowFlags(Qt.WindowCloseButtonHint)

        self.dialog13.setWindowTitle('Scenario13')
        self.dialog13.setWindowModality(Qt.NonModal)
        self.dialog13.show()

    def Dialog14(self):
        self.dialog14 = QDialog()
        self.dialog14.setStyleSheet('background-color: #2E2E38')
        self.dialog14.setWindowIcon(QIcon("./EY_logo.png"))

        cursor = self.cnxn.cursor()

        sql = '''
                 SELECT 											
                        *
                 FROM  [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											

            '''.format(field=self.selected_project_id)

        accountsname = pd.read_sql(sql, self.cnxn)

        self.new_tree = Form(self)

        self.new_tree.tree.clear()

        for n, i in enumerate(accountsname.AccountType.unique()):
            self.new_tree.parent = QTreeWidgetItem(self.new_tree.tree)

            self.new_tree.parent.setText(0, "{}".format(i))
            self.new_tree.parent.setFlags(self.new_tree.parent.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
            child_items = accountsname.AccountSubType[
                accountsname.AccountType == accountsname.AccountType.unique()[n]].unique()
            for m, x in enumerate(child_items):
                self.new_tree.child = QTreeWidgetItem(self.new_tree.parent)

                self.new_tree.child.setText(0, "{}".format(x))
                self.new_tree.child.setFlags(self.new_tree.child.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                grandchild_items = accountsname.AccountClass[accountsname.AccountSubType == child_items[m]].unique()
                for o, y in enumerate(grandchild_items):
                    self.new_tree.grandchild = QTreeWidgetItem(self.new_tree.child)

                    self.new_tree.grandchild.setText(0, "{}".format(y))
                    self.new_tree.grandchild.setFlags(
                        self.new_tree.grandchild.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                    num_name = accountsname[accountsname.AccountClass == grandchild_items[o]].iloc[:, 2:4]
                    full_name = num_name["GLAccountNumber"].map(str) + ' ' + num_name["GLAccountName"]
                    for z in full_name:
                        self.new_tree.grandgrandchild = QTreeWidgetItem(self.new_tree.grandchild)

                        self.new_tree.grandgrandchild.setText(0, "{}".format(z))
                        self.new_tree.grandgrandchild.setFlags(
                            self.new_tree.grandgrandchild.flags() | Qt.ItemIsUserCheckable)
                        self.new_tree.grandgrandchild.setCheckState(0, Qt.Checked)
        self.new_tree.get_selected_leaves()  # 초기값 모두 선택 (추가)

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

        # JE Line Number / JE Number 라디오 버튼
        self.rbtn1 = QRadioButton('JE Line Number', self.dialog14)
        self.rbtn1.setStyleSheet("color: white;")
        font11 = self.rbtn1.font()
        font11.setBold(True)
        self.rbtn1.setFont(font11)
        self.rbtn1.setChecked(True)
        self.rbtn2 = QRadioButton('JE Number', self.dialog14)
        self.rbtn2.setStyleSheet("color: white;")
        font12 = self.rbtn2.font()
        font12.setBold(True)
        self.rbtn2.setFont(font12)

        labelKeyword = QLabel('Key Words* : ', self.dialog14)
        labelKeyword.setStyleSheet("color: white;")

        font1 = labelKeyword.font()
        font1.setBold(True)
        labelKeyword.setFont(font1)

        self.D14_Key = QLineEdit(self.dialog14)
        self.D14_Key.setStyleSheet("background-color: white;")
        self.D14_Key.setPlaceholderText('검색할 단어를 입력하세요')

        labelTE = QLabel('TE : ', self.dialog14)
        labelTE.setStyleSheet("color: white;")

        font2 = labelTE.font()
        font2.setBold(True)
        labelTE.setFont(font2)

        self.D14_TE = QLineEdit(self.dialog14)
        self.D14_TE.setStyleSheet("background-color: white;")
        self.D14_TE.setPlaceholderText('중요성 금액을 입력하세요')

        label_tree = QLabel('특정 계정 : ', self.dialog14)
        label_tree.setStyleSheet("color: white;")
        font4 = label_tree.font()
        font4.setBold(True)
        label_tree.setFont(font4)

        labelSheet = QLabel('시트명* : ', self.dialog14)
        labelSheet.setStyleSheet("color: white;")

        font5 = labelSheet.font()
        font5.setBold(True)
        labelSheet.setFont(font5)

        self.D14_Sheet = QLineEdit(self.dialog14)
        self.D14_Sheet.setStyleSheet("background-color: white;")
        self.D14_Sheet.setPlaceholderText('시트명을 입력하세요')

        self.D14_Key.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D14_TE.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D14_Sheet.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        layout1 = QGridLayout()
        layout1.addWidget(self.rbtn1, 0, 0)
        layout1.addWidget(self.rbtn2, 0, 1)
        layout1.addWidget(labelKeyword, 1, 0)
        layout1.addWidget(self.D14_Key, 1, 1)
        layout1.addWidget(labelTE, 2, 0)
        layout1.addWidget(self.D14_TE, 2, 1)
        layout1.addWidget(label_tree, 3, 0)
        layout1.addWidget(self.new_tree, 3, 1)
        layout1.addWidget(labelSheet, 4, 0)
        layout1.addWidget(self.D14_Sheet, 4, 1)

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

        # ? 제거
        self.dialog14.setWindowFlags(Qt.WindowCloseButtonHint)

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
    def timerEvent(self, e):
        if self.step >= 100:
            self.timer.stop()
            self.btn2.setText('Finished')
            return
        self.step = self.step + 1
        self.pbar.setValue(self.step)

    def doAction(self):
        for i in range(100):
            time.sleep(0.01)
            self.progressBar.setValue(i)

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

        ##시나리오 Sheet를 표현할 콤보박스
        self.combo_sheet = QComboBox(self)

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
        export_file_button.clicked.connect(self.saveFile)
        self.combo_sheet.activated[str].connect(self.Sheet_ComboBox_Selected)

        ##layout 쌓기
        layout = QHBoxLayout()
        layout.addWidget(label_sheet, stretch=1)
        layout.addWidget(self.combo_sheet, stretch=4)
        layout.addWidget(RemoveSheet_button, stretch=1)
        layout.addWidget(export_file_button, stretch=1)
        groupbox.setLayout(layout)

        return groupbox

    def handle_date_clicked(self, date):
        self.dialog6.activateWindow()
        self.D6_Date.setText(date.toString("yyyy-MM-dd"))
        self.dialog6.activateWindow()

    def handle_date_clicked2(self, date):
        self.dialog7.activateWindow()
        self.D7_Date.setText(date.toString("yyyy-MM-dd"))
        self.dialog7.activateWindow()

    def calendar6(self):
        self.dialog6.activateWindow()
        self.new_calendar.show()
        self.dialog6.activateWindow()

    def calendar7(self):
        self.dialog7.activateWindow()
        self.new_calendar.show()
        self.dialog7.activateWindow()

    def extButtonClicked4(self):
        temp_N = self.D4_N.text()
        temp_TE = self.D4_TE.text()
        tempSheet = self.D4_Sheet.text()

        if temp_N == '' or tempSheet == '':
            self.alertbox_open()

        elif self.combo_sheet.findText(tempSheet) != -1: #시트명 중복 확인
            self.alertbox_open5()

        elif checked_account == 'AND JournalEntries.GLAccountNumber IN ()':
            self.alertbox_open6() # 계정 선택 오류

        else:
            cursor = self.cnxn.cursor()

            sql_query = """
            """.format(field=self.selected_project_id)

        self.dataframe = pd.read_sql(sql_query, self.cnxn)

        if len(self.dataframe) > 300000:
            self.alertbox_open3()

        else:
            model = DataFrameModel(self.dataframe)
            self.viewtable.setModel(model)
            self.scenario_dic[tempSheet] = self.dataframe
            key_list = list(self.scenario_dic.keys())
            result = [key_list[0], key_list[-1]]
            self.combo_sheet.addItem(str(result[1]))

    def extButtonClicked5_SAP(self):
        tempSheet_SAP = self.D5_Sheet.text()
        tempYear_SAP = self.D5_Year.text()
        tempSKA1 = self.listbox_drops.text()

        ### 예외처리 1 - 파일 경로 unicode 문제 해결
        dropped_items = []  ### ListBox 인풋값 append
        for i in range(self.listbox_drops.count()):
            dropped_items.append(self.listbox_drops.item(i))

        for i in range(self.dropped_items.count()):
            dropped_items[i] = re.sub(r'\'', '/', dropped_items[i])

        df = pd.DataFrame() ### dataframe으로 저장
        for i in range(len(dropped_items)):
            df = df.append(pd.read_csv(dropped_items[i], sep='|'))

        ### 당기 생성된 계정 코드 반환
        temp_AccCode = list()

        for i in range(len(df)):
            df.loc[i, 'ERDAT'] = str(df.loc[i, 'ERDAT'])
            year = df.loc[i, 'ERDAT'][0:4]

            # ### 당기 시점 지정
            # now = datetime.datetime.now()
            # before_three_months = now - relativedelta(month=3)
            #
            if int(year) == tempYear_SAP:
                temp_AccCode.append(df.loc[i, 'SAKNR'])

        ### 예외처리 2 - 필수값 누락
        if tempYear_SAP == '' or tempSheet_SAP == '' or tempSKA1 == '':
            self.alertbox_open()

        ### 예외처리 3 - 시트명 중복 확인
        elif self.combo_sheet.findText(tempSheet_SAP + '_Result') != -1 or self.combo_sheet.findText(tempSheet_SAP + '_Journals') != -1:
            self.alertbox_open5()

        elif checked_account == 'AND JournalEntries.GLAccountNumber IN ()':
            self.alertbox_open6() # 계정 선택 오류

        else:
            cursor = self.cnxn.cursor()

            sql_query = """""".format(field=self.selected_project_id)

        self.dataframe = pd.read_sql(sql_query, self.cnxn)

        if len(self.dataframe) > 300000:
            self.alertbox_open3()

        else:
            model = DataFrameModel(self.dataframe)
            self.viewtable.setModel(model)
            self.scenario_dic[tempSheet_SAP] = self.dataframe
            key_list = list(self.scenario_dic.keys())
            result = [key_list[0], key_list[-1]]
            self.combo_sheet.addItem(str(result[1]))

    def extButtonClicked5_Non_SAP(self):
        tempSheet_NonSAP = self.D5_Sheet2
        tempYear_NonSAP = self.D5_Year2.text()
        tempCode = self.MyInput.text()

        temp_Code_Non_SAP = self.D5_Code.text()
        temp_Code_Non_SAP = re.sub(r"[:,|\s]", ",", temp_Code_Non_SAP)
        temp_Code_Non_SAP = re.split(",", temp_Code_Non_SAP)

        ### 예외처리 1 - 필수값 입력 누락
        if tempCode == '' or tempSheet_NonSAP == '' or tempYear_NonSAP == '':
            self.alertbox_open()

        elif checked_account == 'AND JournalEntries.GLAccountNumber IN ()':
            self.alertbox_open6() # 계정 선택 오류

        ### 예외처리 2 - 시트명 중복 확인
        elif self.combo_sheet.findText(tempSheet_NonSAP + '_Result') != -1 or self.combo_sheet.findText(tempSheet_NonSAP + '_Journals') != -1:
            self.alertbox_open5()

        else:
            # try:
            #     int(tempYear_NonSAP)

            cursor = self.cnxn.cursor()

            sql_query = """""".format(field=self.selected_project_id)

        self.dataframe = pd.read_sql(sql_query, self.cnxn)

        if len(self.dataframe) > 300000:
            self.alertbox_open3()

        else:
            model = DataFrameModel(self.dataframe)
            self.viewtable.setModel(model)
            self.scenario_dic[tempSheet_NonSAP] = self.dataframe
            key_list = list(self.scenario_dic.keys())
            result = [key_list[0], key_list[-1]]
            self.combo_sheet.addItem(str(result[1]))

    def extButtonClicked6(self):
        tempDate = self.D6_Date.text()
        realDate = date.fromisoformat(tempDate)
        tempTDate = self.D6_Date2.text()
        tempJE = self.D6_JE.text()
        tempCost = self.D6_Cost.text()
        tempSheet = self.D6_Sheet.text()

        if tempDate == '' or tempSheet == '':
            self.alertbox_open()

        elif checked_account == 'AND JournalEntries.GLAccountNumber IN ()':
            self.alertbox_open6() # 계정 선택 오류

        elif self.combo_sheet.findText(tempSheet) != -1: # 시트명 중복 확인
            self.alertbox_open5()

        else:
            if tempCost == '': tempCost = 0
            if tempTDate == '': tempTDate = 0

            checked_account6 = checked_account

            try:
                int(tempTDate)
                int(tempCost)

                first_origin = realDate - timedelta(days=int(tempTDate))
                second_origin = realDate + timedelta(days=int(tempTDate))

                first_mid = str(first_origin).split('-')
                second_mid = str(second_origin).split('-')

                first = "'" + first_mid[0] + first_mid[1] + first_mid[2] + "'"
                second = "'" + second_mid[0] + second_mid[1] + second_mid[2] + "'"

                cursor = self.cnxn.cursor()

                if self.rbtn1.isChecked():
                    sql = '''

                               SELECT											
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
                                   , JournalEntries.Amount											
                                   , JournalEntries.FunctionalCurrencyCode											
                                   , JournalEntries.JEDescription											
                                   , JournalEntries.JELineDescription											
                                   , JournalEntries.PreparerID											
                                   , JournalEntries.ApproverID											
                               FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] JournalEntries,											
                                       [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											
                               WHERE JournalEntries.GLAccountNumber = CoA.GLAccountNumber
                                       AND (JournalEntries.EntryDate BETWEEN {first_date} AND {second_date})
                                       AND ABS(JournalEntries.Amount) > {TE}
                                       AND JournalEntries.PreparerID LIKE '%{tempPrepare}%'
                                       {Account}
                               ORDER BY JournalEntries.JENumber, JournalEntries.JELineNumber

                            '''.format(field=self.selected_project_id, Account=checked_account6, TE=tempCost,
                                       first_date=str(first), second_date=str(second), tempPrepare=tempJE)

                elif self.rbtn2.isChecked():
                    sql = '''

                               SELECT											
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
                                   , JournalEntries.Amount											
                                   , JournalEntries.FunctionalCurrencyCode											
                                   , JournalEntries.JEDescription											
                                   , JournalEntries.JELineDescription											
                                   , JournalEntries.PreparerID											
                                   , JournalEntries.ApproverID											
                               FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] JournalEntries,											
                                       [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											
                               WHERE JournalEntries.GLAccountNumber = CoA.GLAccountNumber AND JournalEntries.JENumber IN		
                                (
                                    SELECT DISTINCT JENumber
                                    FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] AS JournalEntries	
                                    WHERE (JournalEntries.EntryDate BETWEEN {first_date} AND {second_date})
                                            {Account}
                                            AND JournalEntries.PreparerID LIKE '%{tempPrepare}%'
                                            AND ABS(JournalEntries.Amount) > {TE}
                                )
                                ORDER BY JournalEntries.JENumber, JournalEntries.JELineNumber		

                            '''.format(field=self.selected_project_id, Account=checked_account6, TE=tempCost,
                                       first_date=str(first), second_date=str(second), tempPrepare=tempJE)

                self.dataframe = pd.read_sql(sql, self.cnxn)

                if tempJE == '':
                    tempJE = '모두'

                if len(self.dataframe) > 300000:
                    self.alertbox_open3()
                else:
                    model = DataFrameModel(self.dataframe)
                    self.viewtable.setModel(model)
                    self.scenario_dic[tempSheet] = self.dataframe
                    key_list = list(self.scenario_dic.keys())
                    result = [key_list[0], key_list[-1]]
                    self.combo_sheet.addItem(str(result[1]))

                    buttonReply = QMessageBox.information(self, "라인수 추출", "[결산일: " + str(tempDate) + " T일: " + str(
                        tempTDate) + " 중요성금액: " + str(
                        tempCost) + " 전표입력자: " + tempJE + "] 라인수 " + str(len(self.dataframe)) + "개입니다", QMessageBox.Yes)
                    if buttonReply == QMessageBox.Yes:
                        self.dialog6.activateWindow()

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
        tempDate = self.D7_Date.text()
        realDate_first = date.fromisoformat(tempDate)

        realDate_mid = str(realDate_first).split('-')
        realDate = "'" + realDate_mid[0] + realDate_mid[1] + realDate_mid[2] + "'"

        tempJE = self.D7_JE.text()
        tempCost = self.D7_Cost.text()
        tempSheet = self.D7_Sheet.text()

        if self.rbtn1.isChecked():  # Effective Date 일 때
            tempState = 'AND JournalEntries.EffectiveDate IN (' + realDate + ')'

        elif self.rbtn2.isChecked():  # Entry Date 일 때
            tempState = 'AND JournalEntries.EntryDate IN (' + realDate + ')'

        if tempCost == '':
            tempCost = 0

        if tempDate == '' or tempSheet == '':
            self.alertbox_open()

        elif checked_account == 'AND JournalEntries.GLAccountNumber IN ()':
            self.alertbox_open6()

        elif self.combo_sheet.findText(tempSheet) != -1:
            self.alertbox_open5()

        else:
            try:
                int(tempCost)
                checked_account7 = checked_account

                cursor = self.cnxn.cursor()

                if self.rbtn3.isChecked():
                    sql = '''
                               SELECT 											
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
                                    {Date}
                                    {Account}
                                    AND JournalEntries.PreparerID LIKE '%{tempPrepare}%'
                                    AND ABS(JournalEntries.Amount) > {TE}
                               ORDER BY JournalEntries.JENumber, JournalEntries.JELineNumber

                           '''.format(field=self.selected_project_id, TE=tempCost, Date=tempState, tempPrepare=tempJE,
                                      Account=checked_account7)

                elif self.rbtn4.isChecked():
                    sql = '''
                               SELECT 											
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
                                   , JournalEntries.Amount											
                                   , JournalEntries.FunctionalCurrencyCode											
                                   , JournalEntries.JEDescription											
                                   , JournalEntries.JELineDescription											
                                   , JournalEntries.Source											
                                   , JournalEntries.PreparerID											
                                   , JournalEntries.ApproverID											
                               FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] JournalEntries,											
                                       [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											
                               WHERE JournalEntries.GLAccountNumber = CoA.GLAccountNumber AND JournalEntries.JENumber IN (		
                                   SELECT DISTINCT JENumber
                                   FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] AS JournalEntries
                                   WHERE JournalEntries.PreparerID LIKE '%{tempPrepare}%'
                                        {Account}
                                        {Date}
                                        AND ABS(JournalEntries.Amount) > {TE})
                               ORDER BY JournalEntries.JENumber, JournalEntries.JELineNumber


                           '''.format(field=self.selected_project_id, TE=tempCost, Date=tempState, tempPrepare=tempJE,
                                      Account=checked_account7)

                self.dataframe = pd.read_sql(sql, self.cnxn)

                if len(self.dataframe) > 300000:
                    self.alertbox_open3()

                else:
                    model = DataFrameModel(self.dataframe)
                    self.viewtable.setModel(model)
                    self.scenario_dic[tempSheet] = self.dataframe
                    key_list = list(self.scenario_dic.keys())
                    result = [key_list[0], key_list[-1]]
                    self.combo_sheet.addItem(str(result[1]))

                    buttonReply = QMessageBox.information(self, "라인수 추출", "[Effective Date / Entry Date: " + str(
                        tempDate) + " 중요성금액: " + str(
                        tempCost) + " 전표입력자: " + tempJE + "] 라인수 " + str(len(self.dataframe)) + "개입니다", QMessageBox.Yes)
                    if buttonReply == QMessageBox.Yes:
                        self.dialog6.activateWindow()

            except ValueError:
                self.alertbox_open2('중요성 금액')

    def extButtonClicked8(self):
        tempN = self.D8_N.text()
        tempJE = self.D8_JE.text()
        tempCost = self.D8_Cost.text()
        tempSheet = self.D8_Sheet.text()

        realNDate = int(tempN)

        if tempN == '' or tempSheet == '':
            self.alertbox_open()

        elif checked_account == 'AND JournalEntries.GLAccountNumber IN ()':
            self.alertbox_open6()

        elif self.combo_sheet.findText(tempSheet) != -1:
            self.alertbox_open5()

        else:
            if tempCost == '': tempCost = 0
            try:
                int(tempN)
                int(tempCost)

                checked_account8 = checked_account

                cursor = self.cnxn.cursor()

                if self.rbtn1.isChecked():

                    sql = '''
                                    SELECT 
                                       JournalEntries.FileId	
                                        , JournalEntries.BusinessUnit	
                                        , JournalEntries.JENumber	
                                        , JournalEntries.JELineNumber	
                                        , JournalEntries.EffectiveDate	
                                        , JournalEntries.EntryDate	
                                        , JournalEntries.Period	
                                        , JournalEntries.GLAccountNumber	
                                        , CoA.GLAccountName	
                                        , JournalEntries.Debit	
                                        , JournalEntries.Credit	
                                        , JournalEntries.Amount	
                                        , JournalEntries.FunctionalCurrencyCode	
                                        , JournalEntries.JEDescription	
                                        , JournalEntries.JELineDescription	
                                        , JournalEntries.PreparerID	
                                        , JournalEntries.ApproverID	

                                    FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] JournalEntries,											
                                           [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											
                                    WHERE JournalEntries.GLAccountNumber = CoA.GLAccountNumber

                                        AND DATEDIFF(dd, JournalEntries.EntryDate ,JournalEntries.EffectiveDate) >= {realNDate}
                                        AND JournalEntries.PreparerID LIKE '%{tempPrepare}%'
                                        AND ABS(JournalEntries.Amount) > {TE}
                                        {Account}
                                    ORDER BY JournalEntries.JENumber, JournalEntries.JELineNumber
                                '''.format(field=self.selected_project_id, realNDate=realNDate, TE=tempCost,
                                           tempPrepare=tempJE, Account=checked_account8)

                elif self.rbtn2.isChecked():

                    sql = '''
                                    SELECT 
                                       JournalEntries.FileId	
                                        , JournalEntries.BusinessUnit	
                                        , JournalEntries.JENumber	
                                        , JournalEntries.JELineNumber	
                                        , JournalEntries.EffectiveDate	
                                        , JournalEntries.EntryDate	
                                        , JournalEntries.Period	
                                        , JournalEntries.GLAccountNumber	
                                        , CoA.GLAccountName	
                                        , JournalEntries.Debit	
                                        , JournalEntries.Credit	
                                        , JournalEntries.Amount	
                                        , JournalEntries.FunctionalCurrencyCode	
                                        , JournalEntries.JEDescription	
                                        , JournalEntries.JELineDescription	
                                        , JournalEntries.PreparerID	
                                        , JournalEntries.ApproverID	

                                    FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] JournalEntries,											
                                           [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											
                                    WHERE JournalEntries.GLAccountNumber = CoA.GLAccountNumber AND JournalEntries.JENumber IN 

                                        (
                                        SELECT DISTINCT JENumber
                                        FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] AS JournalEntries	
                                        WHERE DATEDIFF(dd, JournalEntries.EntryDate ,JournalEntries.EffectiveDate) >= {realNDate}
                                            AND JournalEntries.PreparerID LIKE '%{tempPrepare}%'
                                            AND ABS(JournalEntries.Amount) > {TE}
                                            {Account}
                                        )
                                    ORDER BY JournalEntries.JENumber, JournalEntries.JELineNumber
                                    '''.format(field=self.selected_project_id, realNDate=realNDate, TE=tempCost,
                                               tempPrepare=tempJE, Account=checked_account8)

                self.dataframe = pd.read_sql(sql, self.cnxn)

                if len(self.dataframe) > 300000:
                    self.alertbox_open3()

                else:
                    model = DataFrameModel(self.dataframe)
                    self.viewtable.setModel(model)
                    self.scenario_dic[tempSheet] = self.dataframe
                    key_list = list(self.scenario_dic.keys())
                    result = [key_list[0], key_list[-1]]
                    self.combo_sheet.addItem(str(result[1]))
                    buttonReply = QMessageBox.information(self, "라인수 추출", "[N일: " + str(tempN) + " 중요성금액: " + str(
                        tempCost) + " 전표입력자: " + str(tempJE) + "] 라인수 " + str(len(self.dataframe)) + "개입니다",
                                                          QMessageBox.Yes)
                    if buttonReply == QMessageBox.Yes:
                        self.dialog8.activateWindow()

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

            tempN = self.D9_N.text()  # 필수값
            tempTE = self.D9_TE.text()
            tempSheet = self.D9_Sheet.text()

            if tempN == '' or tempSheet == '':
                self.alertbox_open()

            elif self.combo_sheet.findText(tempSheet) != -1:  # 시트명 중복 확인
                self.alertbox_open5()

            elif checked_account == 'AND JournalEntries.GLAccountNumber IN ()':
                self.alertbox_open6()  # 계정 선택 오류

            else:
                if tempTE == '': tempTE = 0
                try:
                    int(tempN)
                    int(tempTE)

                    cursor = self.cnxn.cursor()

                    # sql문 수정
                    if self.rbtn1.isChecked():

                        sql = '''
                                   SELECT				
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
    	                                  , JournalEntries.Amount			
    	                                  , JournalEntries.FunctionalCurrencyCode			
    	                                  , JournalEntries.JEDescription			
    	                                  , JournalEntries.JELineDescription			
    	                                  , JournalEntries.PreparerID			
    	                                  , JournalEntries.ApproverID			
                                   FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] AS JournalEntries,				
    	                                  [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] AS CoA			
                                   WHERE JournalEntries.GLAccountNumber = CoA.GLAccountNumber AND JournalEntries.PreparerID IN 				
    	                                  (			
    	                                  SELECT DISTINCT PreparerID			
    	                                  FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries]			
    	                                  GROUP BY PreparerID			
    	                                  HAVING COUNT(GLAccountNumber) <= {N}			
    	                                  )	AND ABS(JournalEntries.Amount) > {TE} {Account}		
                                   ORDER BY JournalEntries.JENumber, JournalEntries.JELineNumber				

                                '''.format(field=self.selected_project_id, TE=tempTE, N=tempN, Account=checked_account)

                        sql_refer = '''
                                   SELECT JournalEntries.PreparerID, COUNT(JournalEntries.PreparerID) AS User_Cnt, SUM(Debit) Sum_of_Debit, SUM(Credit) Sum_of_Credit				
                                   FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] AS JournalEntries				
                                   WHERE JournalEntries.PreparerID IN				
    	                                  (			
    	                                  SELECT DISTINCT PreparerID			
    	                                  FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries]			
    	                                  GROUP BY PreparerID			
    	                                  HAVING COUNT(GLAccountNumber) <= {N}			
    	                                  ) AND ABS(JournalEntries.Amount) > {TE} {Account}			
                                   GROUP BY JournalEntries.PreparerID				

                                '''.format(field=self.selected_project_id, TE=tempTE, N=tempN, Account=checked_account)

                        self.dataframe_refer = pd.read_sql(sql_refer, self.cnxn)


                    elif self.rbtn2.isChecked():

                        sql = '''
                                   SELECT 				
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
    	                               , JournalEntries.Amount			
    	                               , JournalEntries.FunctionalCurrencyCode			
    	                               , JournalEntries.JEDescription			
    	                               , JournalEntries.JELineDescription			
    	                               , JournalEntries.PreparerID			
    	                               , JournalEntries.ApproverID			
                                   FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] AS JournalEntries,				
    	                               [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] AS CoA			
                                   WHERE JournalEntries.GLAccountNumber = CoA.GLAccountNumber AND 				
    	                               JournalEntries.JENumber IN 			
    		                               (		
    		                               SELECT DISTINCT JENumber		
    		                               FROM  [{field}_Import_CY_01].[dbo].[pbcJournalEntries] AS JournalEntries		
    		                               WHERE JournalEntries.PreparerID IN 		
    			                               (	
    			                               SELECT DISTINCT JournalEntries.PreparerID	
    			                               FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] AS JournalEntries,	
    				                               (
    				                               SELECT JournalEntries.PreparerID, COUNT(JournalEntries.PreparerID) AS User_Cnt
    				                               FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] AS JournalEntries
    				                               GROUP BY JournalEntries.PreparerID
    				                               HAVING COUNT(GLAccountNumber) <= {N}
    				                               ) AS LVL1
    			                               WHERE LVL1.PreparerID = JournalEntries.PreparerID	
    			                               ) AND ABS(JournalEntries.Amount) > {TE}	
    		                               ) {Account} 		
                                   ORDER BY JournalEntries.JENumber, JournalEntries.JELineNumber				

                                '''.format(field=self.selected_project_id, TE=tempTE, N=tempN, Account=checked_account)

                    self.dataframe = pd.read_sql(sql, self.cnxn)

                    if len(self.dataframe) > 30000:
                        self.alertbox_open3()

                    elif len(self.dataframe) == 0:
                        self.dataframe = pd.DataFrame({'No Data': ["[전표작성 빈도수: " + str(tempN) + "," + " 중요성금액: " + str(
                            tempTE) + "] 라인수 " + str(len(self.dataframe)) + "개입니다"]})
                        model = DataFrameModel(self.dataframe)
                        model_refer = DataFrameModel(self.dataframe_refer)
                        self.viewtable.setModel(model)

                        if self.rbtn1.isChecked():
                            self.scenario_dic[tempSheet + "_Result"] = self.dataframe
                            self.scenario_dic[tempSheet + "_Reference"] = self.dataframe_refer
                            key_list = list(self.scenario_dic.keys())
                            result = [key_list[0], key_list[-1], key_list[-2]]
                            self.combo_sheet.addItem(str(result[2]))


                        elif self.rbtn2.isChecked():
                            self.scenario_dic[tempSheet + "_Journals"] = self.dataframe
                            key_list = list(self.scenario_dic.keys())
                            result = [key_list[0], key_list[-1]]
                            self.combo_sheet.addItem(str(result[1]))

                        buttonReply = QMessageBox.information(self, "라인수 추출",
                                                              "[전표작성 빈도수: " + str(tempN) + "," + " 중요성금액: " + str(
                                                                  tempTE) + "] 라인수 " + str(
                                                                  len(self.dataframe) - 1) + "개입니다",
                                                              QMessageBox.Yes)
                        if buttonReply == QMessageBox.Yes:
                            self.dialog9.activateWindow()

                    else:

                        if self.rbtn1.isChecked():
                            self.scenario_dic[tempSheet + "_Result"] = self.dataframe
                            self.scenario_dic[tempSheet + "_Reference"] = self.dataframe_refer
                            key_list = list(self.scenario_dic.keys())
                            result = [key_list[0], key_list[-1], key_list[-2]]
                            self.combo_sheet.addItem(str(result[2]))
                            self.combo_sheet.addItem(str(result[1]))
                            model = DataFrameModel(self.dataframe)
                            model_refer = DataFrameModel(self.dataframe_refer)
                            self.viewtable.setModel(model)

                        elif self.rbtn2.isChecked():
                            self.scenario_dic[tempSheet + "_Journals"] = self.dataframe
                            key_list = list(self.scenario_dic.keys())
                            result = [key_list[0], key_list[-1]]
                            self.combo_sheet.addItem(str(result[1]))
                            model = DataFrameModel(self.dataframe)
                            self.viewtable.setModel(model)

                        buttonReply = QMessageBox.information(self, "라인수 추출",
                                                              "[전표작성 빈도수: " + str(tempN) + "," + " 중요성금액: " + str(
                                                                  tempTE) + "] 라인수 " + str(
                                                                  len(self.dataframe)) + "개입니다", QMessageBox.Yes)
                        if buttonReply == QMessageBox.Yes:
                            self.dialog9.activateWindow()

                except ValueError:
                    try:
                        int(tempN)
                        try:
                            int(tempTE)
                        except:
                            self.alertbox_open2('중요성금액')
                    except:
                        try:
                            int(tempTE)
                            self.alertbox_open2('작성빈도수')
                        except:
                            self.alertbox_open2('작성빈도수와 중요성금액')

    def extButtonClicked10(self):
        tempSearch = self.D10_Search.text()  # 필수값
        tempAccount = self.D10_Account.text()
        tempPoint = self.D10_Point.text()
        tempTE = self.D10_TE.text()
        tempSheet = self.D10_Sheet.text()

        if tempSearch == '' or tempAccount == '' or tempPoint == '' or tempTE == '' or tempSheet == '':
            self.alertbox_open()

        elif self.combo_sheet.findText(tempSheet) != -1: #시트명 중복 확인
            self.alertbox_open5()

        elif checked_account == 'AND JournalEntries.GLAccountNumber IN ()':
            self.alertbox_open6() # 계정 선택 오류

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

            if len(self.dataframe) > 300000:
                self.alertbox_open3()

            else:
                model = DataFrameModel(self.dataframe)
                self.viewtable.setModel(model)
                self.scenario_dic[tempSheet] = self.dataframe
                key_list = list(self.scenario_dic.keys())
                result = [key_list[0], key_list[-1]]
                self.combo_sheet.addItem(str(result[1]))

    def extButtonClicked11(self):
        passwords = ''
        users = 'guest'
        server = ids
        password = passwords

        temp_Tree_A = self.account_tree_A.text()
        temp_Tree_B = self.account_tree_B.text()

        if temp_Tree_A == '' or temp_Tree_B == '':
            self.alertbox_open()

        else:
            db = 'master'
            user = users
            cnxn = pyodbc.connect(
                "DRIVER={SQL Server};SERVER=" + server +
                ";uid=" + user +
                ";pwd=" + password +
                ";DATABASE=" + db +
                ";trusted_connection=" + "yes"
            )
            cursor = cnxn.cursor()

            # sql문 수정
            sql_query = '''
                    '''.format(field=fields)

        self.dataframe = pd.read_sql(sql_query, self.cnxn)
        model = DataFrameModel(self.dataframe)
        self.viewtable.setModel(model)

    def extButtonClicked12(self):
        tempCode = self.D12_Code.text()
        tempCost = self.D12_Cost.text()
        tempSheet = self.D12_Sheet.text()

        if self.rbtn1.isChecked():
            tempState = 'Account Name'

        elif self.rbtn2.isChecked():
            tempState = 'Account Code'

        if tempCost == '':
            tempCost = 0

        if tempState == '':
            self.alertbox_open()

        elif checked_account == 'AND JournalEntries.GLAccountNumber IN ()':
            self.alertbox_open6() # 계정 선택 오류

        elif self.combo_sheet.findText(tempSheet) != -1: #시트명 중복 확인
            self.alertbox_open5()

        else:
            try:
                int(tempCode)
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

                self.scenario_dic['특정 계정(A)이 감소할 때 상대계정 리스트 검토_' + str(
                    self.D12_clickcount) + ' (특정 계정(A) = ' + tempCode + ')'] = self.dataframe
                key_list = list(self.scenario_dic.keys())
                result = [key_list[0], key_list[-1]]
                self.combo_sheet.addItem(str(result[1]))

            except ValueError:
                self.alertbox_open2('중요성 금액')

    def extButtonClicked13(self):

        temp_Continuous = self.text_continuous.text()  # 필수
        temp_Tree = self.account_tree.text()
        temp_TE_13 = self.line_amount.text()
        tempSheet = self.D13_Sheet.text()

        if temp_Continuous == '' or temp_TE_13 == '' or temp_Tree == '' or tempSheet == '':
            self.alertbox_open()

        elif self.combo_sheet.findText(tempSheet) != -1: #시트명 중복 확인
            self.alertbox_open5()

        elif checked_account == 'AND JournalEntries.GLAccountNumber IN ()':
            self.alertbox_open6() # 계정 선택 오류

        else:
            cursor = self.cnxn.cursor()

            # sql문 수정
            sql_query = '''
            '''.format(field=self.selected_project_id)

        self.dataframe = pd.read_sql(sql_query, self.cnxn)

        if len(self.dataframe) > 300000:
            self.alertbox_open3()

        else:
            model = DataFrameModel(self.dataframe)
            self.viewtable.setModel(model)
            self.scenario_dic[tempSheet] = self.dataframe
            key_list = list(self.scenario_dic.keys())
            result = [key_list[0], key_list[-1]]
            self.combo_sheet.addItem(str(result[1]))

    def extButtonClicked14(self):

            tempKey = self.D14_Key.text()  # 필수값
            tempTE = self.D14_TE.text()
            tempSheet = self.D14_Sheet.text()

            if tempTE == '' or tempSheet == '':
                self.alertbox_open()

            elif self.combo_sheet.findText(tempSheet) != -1:  # 시트명 중복 확인
                self.alertbox_open5()

            elif checked_account == 'AND JournalEntries.GLAccountNumber IN ()':
                self.alertbox_open6()  # 계정 선택 오류

            else:
                if tempTE == '': tempTE = 0
                try:
                    int(tempTE)

                    cursor = self.cnxn.cursor()

                    # sql 문 수정
                    if self.rbtn1.isChecked():

                        sql = '''
                           SELECT 		
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
    	                        , JournalEntries.Amount	
    	                        , JournalEntries.FunctionalCurrencyCode	
    	                        , JournalEntries.JEDescription	
    	                        , JournalEntries.JELineDescription	
    	                        , JournalEntries.PreparerID	
    	                        , JournalEntries.ApproverID	
                           FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] JournalEntries,		
    	                          [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA	
                           WHERE JournalEntries.GLAccountNumber = CoA.GLAccountNumber 		
    	                          AND (JournalEntries.JEDescription LIKE N'%{KEY}%' OR JournalEntries.JELineDescription LIKE N'%{KEY}%')	
    	                          AND ABS(JournalEntries.Amount) > {TE} {Account}
                           ORDER BY JournalEntries.JENumber, JournalEntries.JELineNumber		

                        '''.format(field=self.selected_project_id, KEY=tempKey, TE=tempTE, Account=checked_account)

                    elif self.rbtn2.isChecked():

                        sql = '''
                           SELECT 			
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
                                , JournalEntries.Amount		
                                , JournalEntries.FunctionalCurrencyCode		
                                , JournalEntries.JEDescription		
                                , JournalEntries.JELineDescription		
                                , JournalEntries.PreparerID		
                                , JournalEntries.ApproverID		
                           FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] JournalEntries,		
    	                       [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA
                           WHERE JournalEntries.GLAccountNumber = CoA.GLAccountNumber AND JournalEntries.JENumber IN			
    	                         (		
                                    SELECT DISTINCT JournalEntries.JENumber		
                                    FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries] AS JournalEntries		
                                    WHERE (JournalEntries.JEDescription LIKE N'%{KEY}%' OR JournalEntries.JELineDescription LIKE N'%{KEY}%')		
                                 )  AND ABS(JournalEntries.Amount) > {TE} {Account}
                           ORDER BY JournalEntries.JENumber, JournalEntries.JELineNumber			
                        '''.format(field=self.selected_project_id, KEY=tempKey, TE=tempTE, Account=checked_account)

                    self.dataframe = pd.read_sql(sql, self.cnxn)
                    if len(self.dataframe) > 300000:
                        self.alertbox_open3()

                    elif len(self.dataframe) == 0:
                        self.dataframe = pd.DataFrame(
                            {'No Data': ["[전표 적요 특정단어: " + str(tempKey) + "," + " 중요성금액: " + str(
                                tempTE) + "] 라인수 " + str(len(self.dataframe)) + "개입니다"]})
                        model = DataFrameModel(self.dataframe)
                        self.viewtable.setModel(model)
                        self.scenario_dic[tempSheet] = self.dataframe
                        key_list = list(self.scenario_dic.keys())
                        result = [key_list[0], key_list[-1]]
                        self.combo_sheet.addItem(str(result[1]))
                        buttonReply = QMessageBox.information(self, "라인수 추출",
                                                              "[전표 적요 특정단어: " + str(tempKey) + "," + " 중요성금액: " + str(
                                                                  tempTE) + "] 라인수 " + str(
                                                                  len(self.dataframe) - 1) + "개입니다",
                                                              QMessageBox.Yes)
                        if buttonReply == QMessageBox.Yes:
                            self.dialog14.activateWindow()

                    else:
                        model = DataFrameModel(self.dataframe)
                        self.viewtable.setModel(model)
                        self.scenario_dic[tempSheet] = self.dataframe
                        key_list = list(self.scenario_dic.keys())
                        result = [key_list[0], key_list[-1]]
                        self.combo_sheet.addItem(str(result[1]))

                        buttonReply = QMessageBox.information(self, "라인수 추출",
                                                              "[전표 적요 특정단어: " + str(tempKey) + "," + " 중요성금액: " + str(
                                                                  tempTE) + "] 라인수 " + str(
                                                                  len(self.dataframe)) + "개입니다", QMessageBox.Yes)
                        if buttonReply == QMessageBox.Yes:
                            self.dialog14.activateWindow()

                except ValueError:
                    try:
                        int(tempTE)
                    except:
                        self.alertbox_open2('중요성금액')

    @pyqtSlot(QModelIndex)
    def slot_clicked_item(self, QModelIndex):
        self.stk_w.setCurrentIndex(QModelIndex.row())

    def saveFile(self):
        if self.dataframe is None:
            self.MessageBox_Open("저장할 데이터가 없습니다")
            return

        if self.scenario_dic == {}:
            self.MessageBox_Open("저장할 Sheet가 없습니다")
            return

        else:
            fileName = QFileDialog.getSaveFileName(self, self.tr("Save Data files"), "./",
                                                   self.tr("xlsx(*.xlsx);; All Files(*.*)"))

            with pd.ExcelWriter('' + fileName[0] + '') as writer:
                for temp in self.scenario_dic:
                    self.scenario_dic[''+temp+''].to_excel(writer, sheet_name= ''+temp+'', index = False, freeze_panes= (1,0))
            self.MessageBox_Open("저장을 완료했습니다")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())
