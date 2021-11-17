import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import pyodbc
import pandas as pd
import numpy as np

class StWidgetForm(QGroupBox):
    """
    위젯 베이스 클래스
    """

    def __init__(self):
        QGroupBox.__init__(self)
        self.box = QBoxLayout(QBoxLayout.TopToBottom)
        self.setLayout(self.box)


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


class MyApp(QWidget):

    def __init__(self):
        super().__init__()
        self.Init_UI()
        ##Initialize Variables
        self.selected_project_id = None
        self.selected_server_name = "서버를 선택하세요"
        self.dataframe = None
        self.users = None
        self.cnxn = None
        self.selected_senario_class = None
        self.selected_senario_subclass = None

    def Init_UI(self):

        img = QImage("C:/Users/BZ297TR/OneDrive - EY/Desktop/helix color.png")
        scaledImg = img.scaled(QSize(1000, 800))

        palette = QPalette()
        palette.setBrush(10, QBrush(scaledImg))
        self.setPalette(palette)

        self.setWindowTitle("TOD ver.2021IntershipProject")

        widget_layout = QBoxLayout(QBoxLayout.TopToBottom)

        self.splitter_layout = QSplitter(Qt.Vertical)

        widget_layout.addSpacing(20)

        self.splitter_layout.addWidget(self.Connect_ServerInfo_Group())
        self.splitter_layout.addWidget(self.Show_DataFrame_Group())
        self.splitter_layout.addWidget(self.Save_Buttons_Group())

        widget_layout.addWidget(self.splitter_layout)
        self.setLayout(widget_layout)

        self.setWindowIcon(QIcon('C:/Users/BZ297TR/OneDrive - EY/Desktop/EY_logo.png'))

        self.setGeometry(150, 150, 1000, 800)
        self.show()

    def Calendar(self):
        cal = QCalendarWidget(self)
        cal.setGridVisible(True)
        cal.clicked[QDate].connect(self.showDate)

        self.lbl = QLabel(self)
        date = cal.selectedDate()
        self.lbl.setText(date.toString())

        vbox = QVBoxLayout()
        vbox.addWidget(cal)
        vbox.addWidget(self.lbl)

        self.setLayout(vbox)

        self.setWindowTitle('Calendar')

        self.setGeometry(150, 150, 1000, 800)
        self.show()

    def showDate(self, date):
        self.lbl.setText(date.toString())

    def Connect_ServerInfo_Group(self):
        groupbox = QGroupBox("접속 정보")
        self.setStyleSheet('QGroupBox {color: white;}')

        layout = QBoxLayout(QBoxLayout.TopToBottom)

        ##SQL 서버 접속 정보
        Login_box = QGridLayout()
        self.label_server = QLabel("Server: ", self)
        self.label_ecode = QLabel("Engagement Code: ", self)
        self.label_pname = QLabel("Project Name: ", self)
        self.label_senario = QLabel("Senario: ", self)

        # 라벨 색상 흰색으로 변경
        self.label_server.setStyleSheet("color: white;")
        self.label_ecode.setStyleSheet("color: white;")
        self.label_pname.setStyleSheet("color: white;")
        self.label_senario.setStyleSheet("color: white;")

        self.combo_server = QComboBox(self)
        self.line_ecode = QLineEdit(self)
        self.combo_pname = QComboBox(self)
        self.combo_senario_class = QComboBox(self)
        self.combo_senario_subclass = QComboBox(self)

        self.combo_server.addItem("서버를 선택하세요")
        for i in range(1, 9):
            self.combo_server.addItem(f'KRSEOVMPPACSQ0{i}\INST1')

        self.combo_pname.addItem("서버가 연결되어 있지 않습니다")

        self.server_connect_button = QPushButton("Connect SQL Server")
        self.server_connect_button.setStyleSheet('color:black;background-color:#FFE602')

        # self.refresh_button = QPushButton("Refresh")
        # self.refresh_button.setStyleSheet('color:black;background-color:#FFE602')

        self.senario_condition_button = QPushButton("Input Conditions")
        self.senario_condition_button.setStyleSheet('color:black;background-color:#FFE602')

        ##SQL 서버를 콤보박스에서 선택
        self.combo_server.activated[str].connect(self.Server_ComboBox_Selected)
        ##SQL Server Connect 버튼을 클릭
        self.server_connect_button.clicked.connect(self.Server_Connect_Button_Clicked)
        ##Project name을 콤보박스에서 선택
        self.combo_pname.activated[str].connect(self.ProjectName_ComboBox_Selected)
        # ##Refresh 버튼을 클릭
        # self.refresh_button.clicked.connect(self.Refresh_Button_Clicked)
        ##시나리오 조건 입력 버튼 클릭
        self.senario_condition_button.clicked.connect(self.Senario_Condition_Button_Clicked)

        Login_box.addWidget(self.label_server, 0, 0)
        Login_box.addWidget(self.combo_server, 0, 1)
        Login_box.addWidget(self.server_connect_button, 0, 2)

        Login_box.addWidget(self.label_ecode, 1, 0)
        Login_box.addWidget(self.line_ecode, 1, 1)

        Login_box.addWidget(self.label_pname, 2, 0)
        Login_box.addWidget(self.combo_pname, 2, 1)

        Login_box.addWidget(self.label_senario, 3, 0)
        Login_box.addWidget(self.combo_senario_class, 3, 1)
        Login_box.addWidget(self.combo_senario_subclass, 4, 1)

        Login_box.addWidget(self.senario_condition_button, 3, 2)
        # Login_box.addWidget(self.refresh_button, 4, 2)

        layout.addLayout(Login_box)
        groupbox.setLayout(layout)

        return groupbox

    def Server_ComboBox_Selected(self, text):
        self.selected_server_name = text

    def Server_Connect_Button_Clicked(self):

        self.users = "guest"

        ecode = self.line_ecode.text()
        server = self.selected_server_name
        password = ""
        db = "master"
        user = self.users

        # 예외처리 - 서버 선택
        if server == "서버를 선택하세요":
            QMessageBox.about(self, "Warning", "서버가 선택되어 있지 않습니다")
            self.Refresh()
            return

        # 예외처리 - 접속 정보 오류
        try:
            self.cnxn = pyodbc.connect("DRIVER={SQL Server};SERVER=" +
                                       server + ";uid=" + user + ";pwd=" +
                                       password + ";DATABASE=" + db +
                                       ";trusted_connection=" + "yes")
        except:
            QMessageBox.about(self, "Warning", "접속 정보가 잘못되었습니다")
            self.Refresh()
            return

        cursor = self.cnxn.cursor()

        sql_query = f"""
                           SELECT ProjectName
                           From [DataAnalyticsRepository].[dbo].[Projects]
                           WHERE EngagementCode IN ({ecode})
                           AND DeletedBy IS NULL

                     """

        # 예외처리 - ecode, pname 오류
        try:
            selected_project_names = pd.read_sql(sql_query, self.cnxn)
        except:
            QMessageBox.about(self, "Warning", "Engagement Code가 잘못되었습니다")
            self.Refresh()
            self.combo_pname.clear()
            self.combo_pname.addItem("프로젝트가 없습니다")
            return

        # 예외처리 - pname 입력 오류
        if len(selected_project_names) == 0:
            QMessageBox.about(self, "Warning", "프로젝트가 없습니다")
            self.Refresh()
            self.combo_pname.clear()
            self.combo_pname.addItem("프로젝트가 없습니다")
            return

        self.combo_pname.clear()
        self.combo_pname.addItem("프로젝트를 선택하세요")
        for i in range(len(selected_project_names)):
            self.combo_pname.addItem(selected_project_names.iloc[i, 0])

        self.Refresh()

    def Refresh(self):
        self.selected_project_id = None
        self.dataframe = None
        self.selected_senario_class = None
        self.selected_senario_subclass = None

        self.viewtable.reset()
        self.viewtable.setModel(DataFrameModel(pd.DataFrame()))
        self.combo_senario_class.clear()
        self.combo_senario_subclass.clear()

    # 버튼 기본 표시값 설정
    # def Refresh_Button_Clicked(self):
    #     self.selected_project_id = None
    #     self.selected_server_name = "서버를 선택하세요"
    #     self.dataframe = None
    #     self.users = None
    #     self.cnxn = None
    #     self.selected_senario_class = None
    #     self.selected_senario_subclass = None
    #
    #     self.line_ecode.setText("")
    #
    #     self.viewtable.reset()
    #     self.viewtable.setModel(DataFrameModel(pd.DataFrame()))
    #     self.combo_server.clear()
    #     self.combo_server.addItem("서버를 선택하세요")
    #     for i in range(1, 9):
    #         self.combo_server.addItem(f'KRSEOVMPPACSQ0{i}\INST1')
    #
    #     self.combo_pname.clear()
    #     self.combo_pname.addItem("서버가 연결되어 있지 않습니다")
    #
    #     self.combo_senario_class.clear()
    #     self.combo_senario_subclass.clear()

    # big, small 변수명 통일해야 함 - Dialog 구성할 때

    def ProjectName_ComboBox_Selected(self, text):

        ecode = self.line_ecode.text()
        pname = text

        if self.cnxn is None:
            return

        cursor = self.cnxn.cursor()

        # 시나리오별 - sql_query 변수명 구분
        sql_query = f"""
                        SELECT Project_ID
                        FROM [DataAnalyticsRepository].[dbo].[Projects]
                        WHERE ProjectName IN (\'{pname}\')
                        AND EngagementCode IN ({ecode})
                        AND DeletedBy is Null

                     """
        # 예외처리 -
        try:
            self.selected_project_id = pd.read_sql(sql_query, self.cnxn).iloc[0, 0]
        except:
            self.selected_project_id = None
            self.Refresh()
            return

        # 시나리오 선택
        self.combo_senario_class.clear()
        self.combo_senario_class.addItem("시나리오를 선택하세요")

        # 코드 짧게 운용 - 보경쌤 코드에서 A팀은 Dialog4
        self.combo_senario_class.addItem('1. Data 완전성', ['(1) 사용 빈도가 N번 이하인 계정이 포함된 전표 리스트 추출',
                                                         '(2) 당기에 생성된 계정 리스트 추출'])
        self.combo_senario_class.addItem('2. Data Timing', ['(1) 결산일 전후 T일 입력 전표 리스트 추출',
                                                            '(2) 비영업일에 전기되거나 입력된 전표 리스트 추출',
                                                            '(3) 효력 일자와 입력 일자 간 차이가 N일 이상인 전표 리스트 추출'])
        self.combo_senario_class.addItem('3. Data 업무분장', ['(1) 전표 작성 빈도수가 N회 이하인 작성자에 의해 생성된 전표 리스트',
                                                          '(2) 특정 작성자에 의해 생성된 전표 리스트'])
        self.combo_senario_class.addItem('4. Data 분개검토', ['(1) 특정 주계정에 대한 상대계정이 매칭되지 않는 전표 리스트 추출',
                                                          '(2) 특정 주계정이 감소할 때의 상대 계정 리스트 추출'])
        self.combo_senario_class.addItem('5. 기타', ['(1) 연속된 숫자로 끝나는 금액을 포함한 전표 리스트 추출',
                                                   '(2) 적요(JE Desription)가 공란이거나 특정 단어(keyword)를 포함하고 있는 전표 리스트 추출'])

        self.combo_senario_class.activated[str].connect(self.SenarioClass_ComboBox_Selected)
        self.combo_senario_subclass.activated[str].connect(self.SenarioSubClass_ComboBox_Selected)

    # 시나리오 세부내역
    def SenarioClass_ComboBox_Selected(self, text):
        self.selected_senario_class = text

        self.combo_senario_subclass.clear()
        self.combo_senario_subclass.addItem("시나리오 세부 내역을 선택하세요")
        index = self.combo_senario_class.currentIndex()

        self.combo_senario_subclass.addItems(self.combo_senario_class.itemData(index))

    def SenarioSubClass_ComboBox_Selected(self, text):
        if self.selected_senario_class is None:
            return

        self.selected_senario_subclass = text

    def Senario_Condition_Button_Clicked(self):
        self.senario_condition_dialog = QDialog()
        self.senario_condition_dialog.setWindowTitle("Input Conditions")
        self.senario_condition_dialog.setWindowModality(Qt.ApplicationModal)

        layout = QVBoxLayout()
        condition_layout = QGridLayout()

        # dialog
        self.label_condition1 = QLabel("조건 1: ", self)
        self.label_condition2 = QLabel("조건 2: ", self)
        self.label_condition3 = QLabel("조건 3: ", self)
        self.label_condition4 = QLabel("조건 4: ", self)

        self.line_condition1 = QLineEdit(self)
        self.line_condition2 = QLineEdit(self)
        self.line_condition3 = QLineEdit(self)
        self.line_condition4 = QLineEdit(self)

        self.line_condition1.setText("")
        self.line_condition2.setText("")
        self.line_condition3.setText("")
        self.line_condition4.setText("")

        condition_layout.addWidget(self.label_condition1, 0, 0)
        condition_layout.addWidget(self.line_condition1, 0, 1)

        condition_layout.addWidget(self.label_condition2, 1, 0)
        condition_layout.addWidget(self.line_condition2, 1, 1)

        condition_layout.addWidget(self.label_condition3, 2, 0)
        condition_layout.addWidget(self.line_condition3, 2, 1)

        condition_layout.addWidget(self.label_condition4, 3, 0)
        condition_layout.addWidget(self.line_condition4, 3, 1)

        button_layout = QHBoxLayout()
        button_layout.setAlignment(Qt.AlignCenter)

        self.dialog_extract_button = QPushButton("Extract Data")

        self.dialog_extract_button.clicked.connect(self.Extract_Button_Clicked)

        button_layout.addWidget(self.dialog_extract_button)

        layout.addSpacing(20)
        layout.addLayout(condition_layout)
        layout.addLayout(button_layout)

        self.senario_condition_dialog.setLayout(layout)
        self.senario_condition_dialog.resize(300, 200)
        self.senario_condition_dialog.show()

    ### Dialog
    def Dialog4(self):
        self.dialog4 = QDialog()
        self.dialog4.setStyleSheet('background-color: #808080')
        self.dialog4.setWindowIcon(QIcon('C:/Users/BZ297TR/OneDrive - EY/Desktop/EY_logo.png'))

        self.btn2 = QPushButton('Data Extract', self.dialog4)
        self.btn2.move(70, 200)
        self.btn2.setStyleSheet('color:black;background-color:#FFE602')
        self.btn2.clicked.connect(self.extButtonClicked4)

        self.btnDialog = QPushButton('Close', self.dialog4)
        self.btnDialog.move(180, 200)
        self.btnDialog.setStyleSheet('color:black;background-color:#FFE602')
        self.btnDialog.clicked.connect(self.dialog_close4)

        label_freq = QLabel('사용빈도(N)* :', self.dialog4)
        label_freq.move(50, 80)
        label_freq.setStyleSheet('color: red;')
        self.D4_N = QLineEdit(self.dialog4)
        self.D4_N.setStyleSheet('background-color: white;')
        self.D4_N.move(150, 80)

        label_TE = QLabel('중요성금액: ', self.dialog4)
        label_TE.move(50, 110)
        label_TE.setStyleSheet('color: white;')
        self.D4_TE = QLineEdit(self.dialog4)
        self.D4_TE.setStyleSheet('background-color: white;')
        self.D4_TE.move(150, 110)

        self.dialog4.setGeometry(300, 300, 350, 300)

        self.dialog4.setWindowTitle('Scenario4')
        self.dialog4.setWindowModality(Qt.ApplicationModal)
        self.dialog4.exec_()

    def Dialog5(self):
        self.dialog5 = QDialog()
        self.dialog5.setStyleSheet('background-color: #808080')
        self.dialog5.setWindowIcon(QIcon('C:/Users/BZ297TR/OneDrive - EY/Desktop/EY_logo.png'))

        self.btn2 = QPushButton('Data Extract', self.dialog5)
        self.btn2.move(70, 200)
        self.btn2.setStyleSheet('color:black;background-color:#FFE602')
        self.btn2.clicked.connect(self.extButtonClicked5)

        self.btnDialog = QPushButton('Close', self.dialog5)
        self.btnDialog.move(180, 200)
        self.btnDialog.setStyleSheet('color:black;background-color:#FFE602')
        self.btnDialog.clicked.connect(self.dialog_close5)

        label_term_start = QLabel('시작일 :', self.dialog5)
        label_term_start.move(50, 80)
        label_term_start.setStyleSheet('color: white;')
        self.D5_term_start = QLineEdit(self.dialog5)
        self.D5_term_start.setInputMask('0000-00-00;*')
        self.D5_term_start.setStyleSheet('background-color: white;')
        self.D5_term_start.move(150, 80)

        label_term_end = QLabel('종료일 :', self.dialog5)
        label_term_end.move(50, 110)
        label_term_end.setStyleSheet('color: white;')
        self.D5_term_end = QLineEdit(self.dialog5)
        self.D5_term_end.setInputMask('0000-00-00;*')
        self.D5_term_end.setStyleSheet('background-color: white;')
        self.D5_term_end.move(150, 110)

        self.dialog5.setGeometry(300, 300, 350, 300)

        self.dialog5.setWindowTitle('Scenario5')
        self.dialog5.setWindowModality(Qt.ApplicationModal)
        self.dialog5.exec_()

    def Dialog13(self):
        self.dialog13 = QDialog()
        self.dialog13.setStyleSheet('background-color: #808080')
        self.dialog13.setWindowIcon(QIcon('C:/Users/BZ297TR/OneDrive - EY/Desktop/EY_logo.png'))

        self.btn2 = QPushButton('Data Extract', self.dialog13)
        self.btn2.move(70, 200)
        self.btn2.setStyleSheet('color:black;background-color:#FFE602')

        self.btnDialog = QPushButton('Close', self.dialog13)
        self.btnDialog.move(180, 200)
        self.btnDialog.setStyleSheet('color:black;background-color:#FFE602')
        self.btnDialog.clicked.connect(self.dialog_close13)

        labelContinuous = QLabel('연속된 자릿수* : ', self.dialog13)
        labelContinuous.move(40, 80)
        labelContinuous.setStyleSheet("color: red;")
        self.D13_continuous = QLineEdit(self.dialog13)
        self.D13_continuous.setStyleSheet("background-color: white;")
        self.D13_continuous.move(170, 80)

        labelAccount = QLabel('특정계정 : ', self.dialog13)
        labelAccount.move(40, 110)
        labelAccount.setStyleSheet("color: white;")
        self.D13_account = QLineEdit(self.dialog13)
        self.D13_account.setStyleSheet("background-color: white;")
        self.D13_account.move(170, 110)

        labelCost = QLabel('중요성금액 : ', self.dialog13)
        labelCost.move(40, 140)
        labelCost.setStyleSheet("color: white;")
        self.D13_cost = QLineEdit(self.dialog13)
        self.D13_cost.setStyleSheet("background-color: white;")
        self.D13_cost.move(170, 140)

        self.dialog13.setGeometry(300, 300, 350, 300)

        self.dialog13.setWindowTitle('Scenario13')
        self.dialog13.setWindowModality(Qt.ApplicationModal)
        self.dialog13.exec_()

    #### extButtonClicked
    def extButtonClicked4(self):
        password = ''
        users = 'guest'
        server = ids  # 서버명
        password = passwords

        temp_N = self.D4_N.text()
        temp_TE = self.D4_TE.text()

        if temp_N == '':
            self.alterbox_open()

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

            sql_query = """

            """.format(field=fields)

        df = pd.read_sql(sql_query, cnxn)

        model = DataFrameModel(df)
        self.viewtable.setModel(model)

        global save_df
        save_df = df

    def extButtonClicked5(self):
        passwords = ''
        users = 'guest'
        server = ids
        password = passwords

        temp_start = self.D5_term_start.text()
        temp_end = self.D5_term_end.text()

        if temp_start == '':
            self.alterbox_open()

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

            sql_query = """

            """.format(field=fields)

        df = pd.read_sql(sql_query, cnxn)

        model = DataFrameModel(df)
        self.viewtable.setModel(model)

        global save_df
        save_df = df

    def extButtonClicked13(self):
        passwords = ''
        users = 'guset'
        server = ids
        password = passwords

        temp_continuous = self.D13_continuous.text()  # 필수
        temp_account_13 = self.D13_account.text()
        temp_cost_13 = self.D13_cost.text()

        if temp_continuous == '' or temp_account_13 == '' or temp_cost_13 == '':
            self.alterbox_open()

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

        df = pd.read_sql(sql_query, cnxn)
        model = DataFrameModel(df)
        self.viewtable.setModel(model)

        global save_df
        save_df = df

    ### Dialog_close
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
        groupbox = QGroupBox('데이터')
        layout = QBoxLayout(QBoxLayout.TopToBottom)

        self.viewtable = QTableView(self)

        layout.addWidget(self.viewtable)
        groupbox.setLayout(layout)

        return groupbox

    # 바로 닫히도록 구현해둠
    def Extract_Button_Clicked(self):

        if self.cnxn is None:
            self.senario_condition_dialog.close()
            return

        project_id = self.selected_project_id

        # 예외처리 - 프로젝트 아이디 공란
        if project_id is None:
            QMessageBox.about(self, "Warning", "프로젝트가 선택되어 있지 않습니다")
            self.senario_condition_dialog.close()
            return

        # condition1 = self.line_condition1.text()
        # condition2 = self.line_condition2.text()
        # condition3 = self.line_condition3.text()
        # condition4 = self.line_condition4.text()

        cursor = self.cnxn.cursor()

        # 임시로 100, 0을 넣어주었다. 나중에 필요에 따라 수정해야 한다.
        sql_query = f"""
                SELECT TOP {100} 
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
               FROM [{project_id}_Import_CY_01].[dbo].[pbcJournalEntries] JournalEntries,											
                       [{project_id}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											
               WHERE JournalEntries.GLAccountNumber = CoA.GLAccountNumber
               AND   ABS(JournalEntries.Amount) >= {0}
               ORDER BY JENumber, JELineNumber
                    """

        self.dataframe = pd.read_sql(sql_query, self.cnxn)
        model = DataFrameModel(self.dataframe)
        self.viewtable.setModel(model)

        self.senario_condition_dialog.close()

    def Save_Buttons_Group(self):

        groupbox = QGroupBox('저장')
        layout1 = QVBoxLayout()
        layout2 = QHBoxLayout()

        self.save_path_button = QPushButton("저장 위치")
        self.save_path_button.setStyleSheet('color:black;background-color:#FFE602')

        self.line_savepath = QLineEdit(self)
        self.export_file_button = QPushButton("저장")
        self.export_file_button.setStyleSheet('color:black;background-color:#FFE602')

        self.save_path_button.clicked.connect(self.SavePathButton_Clicked)
        self.export_file_button.clicked.connect(self.ExportFileButton_Clicked)

        layout2.addWidget(self.save_path_button)
        layout2.addWidget(self.line_savepath)

        layout1.addLayout(layout2)
        layout1.addWidget(self.export_file_button)

        groupbox.setLayout(layout1)

        return groupbox

    @pyqtSlot(QModelIndex)
    def slot_clicked_item(self, QModelIndex):
        self.stk_w.setCurrentIndex(QModelIndex.row())

    # 결과 파일 write
    def SavePathButton_Clicked(self):

        filename = QFileDialog.getSaveFileName(self, "Save File", '',
                                               ".xlsx;; .csv")

        # 예외처리 - 저장 경로 설정
        if filename[0]:
            self.line_path.setText(filename[0] + filename[1])
            self.line_path.setDisabled(True)

        else:
            QMessageBox.about(self, "Warning", "저장 경로를 선택하지 않았습니다.")

    def ExportFileButton_Clicked(self):
        # 예외처리 - 저장 데이터 부재
        if self.dataframe is None:
            QMessageBox.about(self, "Warning", "저장할 데이터가 없습니다")
            return

        save_path = self.line_path.text()

        # 예외처리 - 파일 확장자별 저장 옵션 지정하기
        try:
            if ".xlsx" in save_path:
                self.dataframe.iloc[:, 1:].to_excel(save_path, sheet_name="Sheet1",
                                                    index=False, encoding='utf-8-sig')

            else:  # 저장경로가 .csv 일 경우
                # delimiter |로 저장한다.
                self.dataframe.iloc[:, 1:].to_csv(save_path, sep="|", index=False,
                                                  encoding='utf-8-sig')
            # .csv 선택후 delimiter 지정하는 창 QInpitDialog로 띄우기

        except:
            QMessageBox.about(self, "Warning", "지정된 경로에 이상이 있습니다.")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())
