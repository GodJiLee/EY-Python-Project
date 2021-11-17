import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import numpy as np
import pyodbc
import pandas as pd


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
        if not index.isValid() or not (0 <= index.row() < self.rowCount()
                                       and 0 <= index.column() < self.columnCount()):
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
        self.selected_senario_class_index = None
        self.selected_senario_index = None
        self.password = ""

    ####경고창 함수####
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
        self.alt.setWindowTitle('Warning')
        self.alt.setWindowIcon(QIcon("./EY_logo.png"))
        self.alt.setText('필수 입력값이 누락되었습니다')
        self.alt.exec_()


    ####main_UI####
    def Init_UI(self):

        img = QImage("./dark_gray.png")
        scaledImg = img.scaled(QSize(1000, 900))
        palette = QPalette()
        palette.setBrush(10, QBrush(scaledImg))
        self.setPalette(palette)

        pixmap = QPixmap("./title.png")
        label_img = QLabel()
        label_img.setPixmap(pixmap)
        label_img.setStyleSheet("background-color:#000000")

        widget_layout = QBoxLayout(QBoxLayout.TopToBottom)

        splitter_layout = QSplitter(Qt.Vertical)
        widget_layout.addWidget(label_img)
        splitter_layout.addWidget(self.createFirstExclusiveGroup())
        splitter_layout.addWidget(self.tableview())
        splitter_layout.addWidget(self.createThirdExclusiveGroup())

        widget_layout.addWidget(splitter_layout)

        self.setLayout(widget_layout)

        self.setWindowTitle("Senario Analyzer")
        self.setWindowIcon(QIcon("./EY_logo.png"))
        self.setGeometry(300, 100, 1000, 900)
        self.show()

    def createFirstExclusiveGroup(self):
        groupbox = QGroupBox("접속 정보")

        ##SQL 서버 접속 정보
        grid = QGridLayout()
        label_server = QLabel("Server: ", self)
        label_ecode = QLabel("Engagement Code: ", self)
        label_pname = QLabel("Project Name: ", self)
        label_senario = QLabel("Senario: ", self)

        font_server = label_server.font()
        font_server.setBold(True)
        label_server.setFont(font_server)
        label_server.setStyleSheet("color: white;")

        font_ecode = label_ecode.font()
        font_ecode.setBold(True)
        label_ecode.setFont(font_ecode)
        label_ecode.setStyleSheet("color: white;")

        font_pname = label_pname.font()
        font_pname.setBold(True)
        label_pname.setFont(font_pname)
        label_pname.setStyleSheet("color: white;")

        font_senario = label_senario.font()
        font_senario.setBold(True)
        label_senario.setFont(font_senario)
        label_senario.setStyleSheet("color: white;")

        font_groupbox = groupbox.font()
        font_groupbox.setBold(True)
        groupbox.setFont(font_groupbox)
        self.setStyleSheet('QGroupBox  {color: white;}')

        self.combo_server = QComboBox(self)
        self.line_ecode = QLineEdit(self)
        self.line_ecode.setText("")

        self.combo_pname = QComboBox(self)
        self.combo_senario_class = QComboBox(self)
        self.combo_senario_subclass = QComboBox(self)

        self.combo_server.addItem("서버를 선택하세요")
        for i in range(1, 9):
            self.combo_server.addItem(f'KRSEOVMPPACSQ0{i}\INST1')
        self.combo_server.move(50, 50)

        self.combo_pname.addItem("서버가 연결되어 있지 않습니다")

        self.combo_senario_class.addItem("시나리오를 선택해주세요")

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

        server_connect_button = QPushButton("   SQL Server Connect", self)
        server_connect_button.setStyleSheet('color:white;  background-image : url(./bar.png)')
        font_connect_button = server_connect_button.font()
        font_connect_button.setBold(True)
        server_connect_button.setFont(font_connect_button)


        senario_condition_button = QPushButton("   Input Conditions", self)
        senario_condition_button.setStyleSheet('color:white;  background-image : url(./bar.png)')
        font_condition_button = senario_condition_button.font()
        font_condition_button.setBold(True)
        senario_condition_button.setFont(font_condition_button)

        ##SQL 서버를 콤보박스에서 선택
        self.combo_server.activated[str].connect(self.Server_ComboBox_Selected)
        ##SQL Server Connect 버튼을 클릭
        server_connect_button.clicked.connect(self.SQL_Connect_Button_Clicked)
        ##Project name을 콤보박스에서 선택
        self.combo_pname.activated[str].connect(self.ProjectName_ComboBox_Selected)
        ##시나리오 조건 입력 버튼 클릭
        senario_condition_button.clicked.connect(self.Senario_Condition_Button_Clicked)
        ##첫번째 시나리오를 콤보박스에서 선택
        self.combo_senario_class.activated[str].connect(self.SenarioClass_ComboBox_Selected)
        ##두번째 시나리오를 콤보박스에서 선택
        self.combo_senario_subclass.activated[str].connect(self.SenarioSubClass_ComboBox_Selected)

        grid.addWidget(label_server, 0, 0)
        grid.addWidget(self.combo_server, 0, 1)
        grid.addWidget(server_connect_button, 0, 2)

        grid.addWidget(label_ecode, 1, 0)
        grid.addWidget(self.line_ecode, 1, 1)

        grid.addWidget(label_pname, 2, 0)
        grid.addWidget(self.combo_pname, 2, 1)

        grid.addWidget(label_senario, 3, 0)
        grid.addWidget(self.combo_senario_class, 3, 1)
        grid.addWidget(self.combo_senario_subclass, 4, 1)

        grid.addWidget(senario_condition_button, 3, 2)

        groupbox.setLayout(grid)

        return groupbox

    def tableview(self):
        groupbox = QGroupBox("데이터 추출")
        self.setStyleSheet('QGroupBox  {color: white;}')
        font_groupbox = groupbox.font()
        font_groupbox.setBold(True)
        groupbox.setFont(font_groupbox)

        layout = QBoxLayout(QBoxLayout.TopToBottom)

        self.viewtable = QTableView(self)

        layout.addWidget(self.viewtable)
        groupbox.setLayout(layout)

        return groupbox

    def createThirdExclusiveGroup(self):

        groupbox = QGroupBox("저장")
        font_groupbox = groupbox.font()
        font_groupbox.setBold(True)
        groupbox.setFont(font_groupbox)
        self.setStyleSheet('QGroupBox  {color: white;}')

        layout1 = QVBoxLayout()
        layout2 = QHBoxLayout()
        layout3 = QHBoxLayout()

        label_savepath = QLabel("Save Route: ", self)
        font_savepath = label_savepath.font()
        font_savepath.setBold(True)
        label_savepath.setFont(font_savepath)
        label_savepath.setStyleSheet('color:white;')

        save_path_button = QPushButton("    Setting Save Route", self)
        font_path_button = save_path_button.font()
        font_path_button.setBold(True)
        save_path_button.setFont(font_path_button)
        save_path_button.setStyleSheet('color:white;background-image : url(./bar.png)')

        self.line_savepath = QLineEdit(self)
        self.line_savepath.setText("")
        self.line_savepath.setDisabled(True)

        export_file_button = QPushButton("    Save", self)
        font_export_button = export_file_button.font()
        font_export_button.setBold(True)
        export_file_button.setFont(font_export_button)
        export_file_button.setStyleSheet('color:white;background-image : url(./bar.png)')


        save_path_button.clicked.connect(self.SavePathButton_Clicked)
        export_file_button.clicked.connect(self.ExportFileButton_Clicked)

        layout2.addWidget(label_savepath)
        layout2.addWidget(self.line_savepath)

        layout3.addStretch(4)
        layout3.addWidget(save_path_button)
        layout3.addWidget(export_file_button)

        layout1.addLayout(layout2)
        layout1.addLayout(layout3)

        groupbox.setLayout(layout1)

        return groupbox

    ####ComboBox 선택 값 함수####
    def Server_ComboBox_Selected(self, text):
        self.selected_server_name = text

    def ProjectName_ComboBox_Selected(self, text):

        ecode = self.line_ecode.text()
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

        try:
            self.selected_project_id = pd.read_sql(sql_query, self.cnxn).iloc[0, 0]
        except:
            self.selected_project_id = None
            return

    def SenarioClass_ComboBox_Selected(self, text):
        self.selected_senario_class_index = self.combo_senario_class.currentIndex()

        index = self.combo_senario_class.currentIndex()
        self.combo_senario_subclass.clear()
        ##시나리오가 선택되지 않았을 때
        if index == 0:
            self.selected_senario_class_index = None
            self.selected_senario_subclass_index = None
            return
        self.combo_senario_subclass.addItem("시나리오 세부 내역을 선택하세요")

        self.combo_senario_subclass.addItems(self.combo_senario_class.itemData(index))

    def SenarioSubClass_ComboBox_Selected(self, text):

        if self.selected_senario_class_index is None:
            return

        self.selected_senario_subclass_index = self.combo_senario_subclass.currentIndex()

    ####SQL Connect 함수####

    def SQL_Server_Dialog_Close(self):
        self.server_connect_dialog.close()

    def SQL_Connect_Button_Clicked(self):

        ## 예외처리 - 서버 선택이 되지 않았을 시
        if self.selected_server_name == "서버를 선택하세요":
            self.MessageBox_Open("서버가 선택되어 있지 않습니다")
            return

        ## 예외처리 - Engagement 코드에 문자열이 입력되었을 시
        elif self.line_ecode.text().isdigit() is False:
            self.MessageBox_Open("Engagement Code가 잘못되었습니다")
            return

        self.server_connect_dialog = QDialog()
        # self.server_connect_dialog.setStyleSheet('background-color: #808080')
        self.server_connect_dialog.setWindowIcon(QIcon('./EY_logo.png'))
        self.server_connect_dialog.setWindowTitle("Connect SQL Server")
        self.server_connect_dialog.setWindowModality(Qt.NonModal)

        label_password = QLabel("Password: ", self.server_connect_dialog)
        # label_password.setStyleSheet("color: black;")
        font_password = label_password.font()
        font_password.setBold(True)
        label_password.setFont(font_password)
        label_password.move(30, 70)

        self.line_password = QLineEdit(self.server_connect_dialog)
        self.line_password.setText("")
        self.line_password.setEchoMode(QLineEdit.Password)
        # self.line_password.setStyleSheet("background-color: white;")
        self.line_password.move(120, 70)

        connect_button = QPushButton("Connect", self.server_connect_dialog)
        # connect_button.setStyleSheet('color:black;background-color:#FFE602')
        font_connect_button = connect_button.font()
        font_connect_button.setBold(True)
        connect_button.setFont(font_connect_button)
        connect_button.move(60, 150)

        close_button = QPushButton("Close", self.server_connect_dialog)
        # close_button.setStyleSheet('color:black;background-color:#FFE602')
        font_close_button = close_button.font()
        font_close_button.setBold(True)
        close_button.setFont(font_close_button)
        close_button.move(170, 150)

        connect_button.clicked.connect(self.Connect_Button_Clicked_In_Dialog)
        close_button.clicked.connect(self.SQL_Server_Dialog_Close)

        self.server_connect_dialog.resize(300, 200)
        self.server_connect_dialog.exec_()

    def Connect_Button_Clicked_In_Dialog(self):

        self.users = "guest"
        self.password = self.line_password.text()
        ecode = self.line_ecode.text()
        server = self.selected_server_name
        db = "master"
        user = self.users
        password = self.password

        ## 예외처리 - 패스워드 오류(아마 이 케이스는 일어나지는 않을 것이다)
        try:
            self.cnxn = pyodbc.connect("DRIVER={SQL Server};SERVER=" +
                                       server + ";uid=" + user + ";pwd=" +
                                       password + ";DATABASE=" + db +
                                       ";trusted_connection=" + "yes")
        except:
            self.MessageBox_Open("패스워드가 잘못되었습니다")
            self.line_password.setText("")
            return

        cursor = self.cnxn.cursor()

        sql_query = f"""
                           SELECT ProjectName
                           From [DataAnalyticsRepository].[dbo].[Projects]
                           WHERE EngagementCode IN ({ecode})
                           AND DeletedBy IS NULL
                     """

        selected_project_names = pd.read_sql(sql_query, self.cnxn)

        ## 예외처리 - 해당 Engagement code에 프로젝트가 없는 경우
        if len(selected_project_names) == 0:
            self.MessageBox_Open("프로젝트가 없습니다")
            self.combo_pname.clear()
            self.combo_pname.addItem("해당 Engagement Code에는 프로젝트가 없습니다")
            self.server_connect_dialog.close()
            return

        self.selected_project_id = None
        self.combo_pname.clear()
        self.combo_pname.addItem("프로젝트를 선택하세요")
        for i in range(len(selected_project_names)):
            self.combo_pname.addItem(selected_project_names.iloc[i, 0])

        self.server_connect_dialog.close()

    ####Local로 저장하는 함수들####

    def SavePathButton_Clicked(self):

        filename = QFileDialog.getSaveFileName(self, "Save File", '',
                                               ".xlsx;; .csv")

        if filename[0]:
            self.line_savepath.setText(filename[0] + filename[1])
            self.line_savepath.setDisabled(True)

        else:
            self.MessageBox_Open("저장 경로를 선택하지 않았습니다.")

    def ExportFileButton_Clicked(self):

        if self.dataframe is None:
            self.MessageBox_Open("저장할 데이터가 없습니다")
            return

        save_path = self.line_savepath.text()

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
            self.MessageBox_Open("지정된 경로에 이상이 있습니다.")

    ####Input Conditions 함수####
    def Senario_Condition_Button_Clicked(self):

        if self.cnxn is None:
            self.MessageBox_Open("SQL 서버가 연결되어 있지 않습니다")
            return

        elif self.selected_project_id is None:
            self.MessageBox_Open("프로젝트가 없습니다")
            return

        elif self.selected_senario_class_index is None or self.selected_senario_subclass_index is None:
            self.MessageBox_Open("시나리오가 없습니다")
            return

        elif self.combo_senario_class.currentIndex() == 0 or self.combo_senario_subclass.currentIndex() == 0:
            self.MessageBox_Open("시나리오가 없습니다")
            return

        if self.selected_senario_class_index == 1 and self.selected_senario_subclass_index == 1:
            self.Senario04_Dialog()

        elif self.selected_senario_class_index == 1 and self.selected_senario_subclass_index == 2:
            self.Senario05_Dialog()

        elif self.selected_senario_class_index == 2 and self.selected_senario_subclass_index == 1:
            self.Senario06_Dialog()

        elif self.selected_senario_class_index == 2 and self.selected_senario_subclass_index == 2:
            self.Senario07_Dialog()

        elif self.selected_senario_class_index == 2 and self.selected_senario_subclass_index == 3:
            self.Senario08_Dialog()

        elif self.selected_senario_class_index == 3 and self.selected_senario_subclass_index == 1:
            self.Senario09_Dialog()

        elif self.selected_senario_class_index == 3 and self.selected_senario_subclass_index == 2:
            self.Senario10_Dialog()

        elif self.selected_senario_class_index == 4 and self.selected_senario_subclass_index == 1:
            self.Senario11_Dialog()

        elif self.selected_senario_class_index == 4 and self.selected_senario_subclass_index == 2:
            self.Senario12_Dialog()

        elif self.selected_senario_class_index == 5 and self.selected_senario_subclass_index == 1:
            self.Senario13_Dialog()

        elif self.selected_senario_class_index == 5 and self.selected_senario_subclass_index == 2:
            self.Senario14_Dialog()

    @pyqtSlot(QModelIndex)
    def slot_clicked_item(self, QModelIndex):
        self.stk_w.setCurrentIndex(QModelIndex.row())

    ####Input Conditions관련 Dialog함수####

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

    def Senario04_Dialog(self):
        self.dialog4 = QDialog()
        self.dialog4.setStyleSheet('background-color: #2E2E38')
        self.dialog4.setWindowIcon(QIcon('./EY_logo.png'))

        self.btn2 = QPushButton('Data Extract', self.dialog4)
        self.btn2.move(70, 200)
        self.btn2.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked4)

        self.btnDialog = QPushButton('Close', self.dialog4)
        self.btnDialog.move(180, 200)
        self.btnDialog.setStyleSheet('color:white;  background-image : url(./bar.png)')
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
        self.dialog4.setWindowModality(Qt.NonModal)
        self.dialog4.show()

    def Senario05_Dialog(self):
        self.dialog5 = QDialog()
        self.dialog5.setStyleSheet('background-color: #2E2E38')
        self.dialog5.setWindowIcon(QIcon('./EY_logo.png'))

        self.btn2 = QPushButton('Data Extract', self.dialog5)
        self.btn2.move(70, 200)
        self.btn2.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked5)

        self.btnDialog = QPushButton('Close', self.dialog5)
        self.btnDialog.move(180, 200)
        self.btnDialog.setStyleSheet('color:white;  background-image : url(./bar.png)')
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
        self.dialog5.setWindowModality(Qt.NonModal)
        self.dialog5.show()


    def Senario06_Dialog(self):
        self.dialog6 = QDialog()
        self.dialog6.setStyleSheet('background-color: #2E2E38')
        self.dialog6.setWindowIcon(QIcon("./EY_logo.png"))

        self.btn2 = QPushButton(' Extract Data', self.dialog6)
        self.btn2.move(70, 200)
        self.btn2.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked6)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        self.btnDialog = QPushButton(" Close", self.dialog6)
        self.btnDialog.move(180, 200)
        self.btnDialog.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close6)

        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)

        labelDate = QLabel('결산일* : ', self.dialog6)
        labelDate.move(50, 50)
        labelDate.setStyleSheet("color: white;")

        font1 = labelDate.font()
        font1.setBold(True)
        labelDate.setFont(font1)

        self.D6_Date = QLineEdit(self.dialog6)
        self.D6_Date.setStyleSheet("background-color: white;")
        self.D6_Date.move(150, 50)
        self.D6_Date.setInputMask("0000-00-00;*")

        labelDate2 = QLabel('T일* : ', self.dialog6)
        labelDate2.move(50, 80)
        labelDate2.setStyleSheet("color: white;")

        font2 = labelDate2.font()
        font2.setBold(True)
        labelDate2.setFont(font2)

        self.D6_Date2 = QLineEdit(self.dialog6)
        self.D6_Date2.setStyleSheet("background-color: white;")
        self.D6_Date2.move(150, 80)

        labelAccount = QLabel('특정계정 : ', self.dialog6)
        labelAccount.move(50, 110)
        labelAccount.setStyleSheet("color: white;")

        font3 = labelAccount.font()
        font3.setBold(True)
        labelAccount.setFont(font3)

        self.D6_Account = QLineEdit(self.dialog6)
        self.D6_Account.setStyleSheet("background-color: white;")
        self.D6_Account.move(150, 110)

        labelJE = QLabel('전표입력자 : ', self.dialog6)
        labelJE.move(50, 140)
        labelJE.setStyleSheet("color: white;")

        font4 = labelJE.font()
        font4.setBold(True)
        labelJE.setFont(font4)

        self.D6_JE = QLineEdit(self.dialog6)
        self.D6_JE.setStyleSheet("background-color: white;")
        self.D6_JE.move(150, 140)

        labelCost = QLabel('중요성금액 : ', self.dialog6)
        labelCost.move(50, 170)
        labelCost.setStyleSheet("color: white;")

        font5 = labelCost.font()
        font5.setBold(True)
        labelCost.setFont(font5)

        self.D6_Cost = QLineEdit(self.dialog6)
        self.D6_Cost.setStyleSheet("background-color: white;")
        self.D6_Cost.move(150, 170)

        self.dialog6.setGeometry(300, 300, 350, 300)

        self.dialog6.setWindowTitle("Scenario6")
        self.dialog6.setWindowModality(Qt.NonModal)
        self.dialog6.show()


    def Senario07_Dialog(self):
        self.dialog7 = QDialog()
        self.dialog7.setStyleSheet('background-color: #2E2E38')
        self.dialog7.setWindowIcon(QIcon("./EY_logo.png"))

        self.btn2 = QPushButton(' Extract Data', self.dialog7)
        self.btn2.move(80, 200)
        self.btn2.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked7)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        self.btnDialog = QPushButton(" Close", self.dialog7)
        self.btnDialog.move(190, 200)
        self.btnDialog.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close7)

        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)

        self.rbtn1 = QRadioButton('Effective Date', self.dialog7)
        self.rbtn1.move(10, 50)
        self.rbtn1.setStyleSheet("color: white;")

        font1 = self.rbtn1.font()
        font1.setBold(True)
        self.rbtn1.setFont(font1)

        self.rbtn1.setChecked(True)

        self.rbtn2 = QRadioButton('Entry Date', self.dialog7)
        self.rbtn2.setStyleSheet("color: white;")
        self.rbtn2.move(200, 50)

        font2 = self.rbtn2.font()
        font2.setBold(True)
        self.rbtn2.setFont(font2)

        labelDate = QLabel('Effective Date/Entry Date* : ', self.dialog7)
        labelDate.move(10, 80)
        labelDate.setStyleSheet("color: white;")

        font3 = labelDate.font()
        font3.setBold(True)
        labelDate.setFont(font3)

        self.D7_Date = QLineEdit(self.dialog7)
        self.D7_Date.setInputMask("0000-00-00;*")
        self.D7_Date.setStyleSheet("background-color: white;")
        self.D7_Date.move(200, 80)

        labelAccount = QLabel('특정계정 : ', self.dialog7)
        labelAccount.move(10, 110)
        labelAccount.setStyleSheet("color: white;")

        font4 = labelAccount.font()
        font4.setBold(True)
        labelAccount.setFont(font4)

        self.D7_Account = QLineEdit(self.dialog7)
        self.D7_Account.setStyleSheet("background-color: white;")
        self.D7_Account.move(200, 110)

        labelJE = QLabel('전표입력자 : ', self.dialog7)
        labelJE.move(10, 140)
        labelJE.setStyleSheet("color: white;")

        font5 = labelJE.font()
        font5.setBold(True)
        labelJE.setFont(font5)

        self.D7_JE = QLineEdit(self.dialog7)
        self.D7_JE.setStyleSheet("background-color: white;")
        self.D7_JE.move(200, 140)

        labelCost = QLabel('중요성금액 : ', self.dialog7)
        labelCost.move(10, 170)
        labelCost.setStyleSheet("color: white;")

        font6 = labelCost.font()
        font6.setBold(True)
        labelCost.setFont(font6)

        self.D7_Cost = QLineEdit(self.dialog7)
        self.D7_Cost.setStyleSheet("background-color: white;")
        self.D7_Cost.move(200, 170)

        self.dialog7.setGeometry(300, 300, 350, 300)

        self.dialog7.setWindowTitle("Scenario7")
        self.dialog7.setWindowModality(Qt.NonModal)
        self.dialog7.show()

    def Senario08_Dialog(self):
        self.dialog8 = QDialog()
        self.dialog8.setStyleSheet('background-color: #2E2E38')
        self.dialog8.setWindowIcon(QIcon("./EY_logo.png"))

        self.btn2 = QPushButton(' Extract Data', self.dialog8)
        self.btn2.move(60, 180)
        self.btn2.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked8)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        self.btnDialog = QPushButton(" Close", self.dialog8)
        self.btnDialog.move(170, 180)
        self.btnDialog.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close8)

        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)

        labelDate = QLabel('N일* : ', self.dialog8)
        labelDate.move(50, 50)
        labelDate.setStyleSheet("color: white;")

        font1 = labelDate.font()
        font1.setBold(True)
        labelDate.setFont(font1)

        self.D8_N = QLineEdit(self.dialog8)
        self.D8_N.setStyleSheet("background-color: white;")
        self.D8_N.move(150, 50)

        labelAccount = QLabel('특정계정 : ', self.dialog8)
        labelAccount.move(50, 80)
        labelAccount.setStyleSheet("color: white;")

        font2 = labelAccount.font()
        font2.setBold(True)
        labelAccount.setFont(font2)

        self.D8_Account = QLineEdit(self.dialog8)
        self.D8_Account.setStyleSheet("background-color: white;")
        self.D8_Account.move(150, 80)

        labelJE = QLabel('전표입력자 : ', self.dialog8)
        labelJE.move(50, 110)
        labelJE.setStyleSheet("color: white;")

        font3 = labelJE.font()
        font3.setBold(True)
        labelJE.setFont(font3)

        self.D8_JE = QLineEdit(self.dialog8)
        self.D8_JE.setStyleSheet("background-color: white;")
        self.D8_JE.move(150, 110)

        labelCost = QLabel('중요성금액 : ', self.dialog8)
        labelCost.move(50, 140)
        labelCost.setStyleSheet("color: white;")

        font4 = labelCost.font()
        font4.setBold(True)
        labelCost.setFont(font4)

        self.D8_Cost = QLineEdit(self.dialog8)
        self.D8_Cost.setStyleSheet("background-color: white;")
        self.D8_Cost.move(150, 140)

        self.dialog8.setGeometry(300, 300, 350, 300)

        self.dialog8.setWindowTitle("Scenario8")
        self.dialog8.setWindowModality(Qt.NonModal)
        self.dialog8.show()


    def Senario09_Dialog(self):
        self.dialog9 = QDialog()
        self.dialog9.setStyleSheet('background-color: #808080')
        self.dialog9.setWindowIcon(QIcon('C:/Users/BZ297TR/OneDrive - EY/Desktop/EY_logo.png'))

        self.btn2 = QPushButton('Data Extract', self.dialog9)
        self.btn2.move(60, 180)
        self.btn2.setStyleSheet('color:black;background-color:#FFE602')
        self.btn2.clicked.connect(self.extButtonClicked9)

        self.btnDialog = QPushButton("닫기", self.dialog9)
        self.btnDialog.move(170, 180)
        self.btnDialog.setStyleSheet('color:black;background-color:#FFE602')
        self.btnDialog.clicked.connect(self.dialog_close9)

        labelKeyword = QLabel('작성빈도수 : ', self.dialog9)
        labelKeyword.move(50, 50)
        labelKeyword.setStyleSheet("color: white;")
        self.D9_N = QLineEdit(self.dialog9)
        self.D9_N.setStyleSheet("background-color: white;")
        self.D9_N.move(150, 50)

        labelTE = QLabel('TE : ', self.dialog9)
        labelTE.move(50, 80)
        labelTE.setStyleSheet("color: white;")
        self.D9_TE = QLineEdit(self.dialog9)
        self.D9_TE.setStyleSheet("background-color: white;")
        self.D9_TE.move(150, 80)

        self.dialog9.setWindowTitle("Scenario9")
        self.dialog9.setWindowModality(Qt.NonModal)
        self.dialog9.show()

    def Senario10_Dialog(self):
        self.dialog10 = QDialog()
        self.dialog10.setStyleSheet('background-color: #808080')
        self.dialog10.setWindowIcon(QIcon('C:/Users/BZ297TR/OneDrive - EY/Desktop/EY_logo.png'))

        self.btn2 = QPushButton('Data Extract', self.dialog10)
        self.btn2.move(60, 180)
        self.btn2.setStyleSheet('color:black;background-color:#FFE602')
        self.btn2.clicked.connect(self.extButtonClicked10)

        self.btnDialog = QPushButton("닫기", self.dialog10)
        self.btnDialog.move(170, 180)
        self.btnDialog.setStyleSheet('color:black;background-color:#FFE602')
        self.btnDialog.clicked.connect(self.dialog_close10)

        labelKeyword = QLabel('전표입력자 : ', self.dialog10)
        labelKeyword.move(50, 30)
        labelKeyword.setStyleSheet("color: white;")
        self.D10_Search = QLineEdit(self.dialog10)
        self.D10_Search.setStyleSheet("background-color: white;")
        self.D10_Search.move(150, 30)

        labelPoint = QLabel('특정시점 : ', self.dialog10)
        labelPoint.move(50, 60)
        labelPoint.setStyleSheet("color: white;")
        self.D10_Point = QLineEdit(self.dialog10)
        self.D10_Point.setStyleSheet("background-color: white;")
        self.D10_Point.move(150, 60)

        labelAccount = QLabel('특정계정 : ', self.dialog10)
        labelAccount.move(50, 90)
        labelAccount.setStyleSheet("color: white;")
        self.D10_Account = QLineEdit(self.dialog10)
        self.D10_Account.setStyleSheet("background-color: white;")
        self.D10_Account.move(150, 90)

        labelTE = QLabel('TE : ', self.dialog10)
        labelTE.move(50, 120)
        self.D10_TE = QLineEdit(self.dialog10)
        self.D10_TE.setStyleSheet("background-color: white;")
        self.D10_TE.move(150, 120)

        self.dialog10.setWindowTitle("Scenario10")
        self.dialog10.setWindowModality(Qt.NonModal)
        self.dialog10.show()

    # A,B,C조 작성 - 추후 논의
    def Senario11_Dialog(self):
        return

    # A,B,C조 작성 - 추후 논의
    def Senario12_Dialog(self):
        return

    def Senario13_Dialog(self):
        self.dialog13 = QDialog()
        self.dialog13.setStyleSheet('background-color: #2E2E38')
        self.dialog13.setWindowIcon(QIcon('./EY_logo.png'))

        self.btn2 = QPushButton('Data Extract', self.dialog13)
        self.btn2.move(70, 200)
        self.btn2.setStyleSheet('color:white;  background-image : url(./bar.png')
        self.btn2.clicked.connect(self.extButtonClicked13)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        self.btnDialog = QPushButton('Close', self.dialog13)
        self.btnDialog.move(180, 200)
        self.btnDialog.setStyleSheet('color:white;  background-image : url(./bar.png)')
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
        self.dialog13.setWindowModality(Qt.NonModal)
        self.dialog13.show()

    def Senario14_Dialog(self):
        self.dialog14 = QDialog()
        self.dialog14.setStyleSheet('background-color: #808080')
        self.dialog14.setWindowIcon(QIcon('C:/Users/BZ297TR/OneDrive - EY/Desktop/EY_logo.png'))

        self.btn2 = QPushButton('Data Extract', self.dialog14)
        self.btn2.move(60, 180)
        self.btn2.setStyleSheet('color:black;background-color:#FFE602')
        self.btn2.clicked.connect(self.extButtonClicked14)

        self.btnDialog = QPushButton("닫기", self.dialog14)
        self.btnDialog.move(170, 180)
        self.btnDialog.setStyleSheet('color:black;background-color:#FFE602')
        self.btnDialog.clicked.connect(self.dialog_close14)

        labelKeyword = QLabel('Key Words : ', self.dialog14)
        labelKeyword.move(50, 50)
        labelKeyword.setStyleSheet("color: white;")
        self.D14_Key = QLineEdit(self.dialog14)
        self.D14_Key.setStyleSheet("background-color: white;")
        self.D14_Key.move(150, 50)

        labelTE = QLabel('TE : ', self.dialog14)
        labelTE.move(50, 80)
        labelTE.setStyleSheet("color: white;")
        self.D14_TE = QLineEdit(self.dialog14)
        self.D14_TE.setStyleSheet("background-color: white;")
        self.D14_TE.move(150, 80)

        self.dialog14.setWindowTitle("Scenario14")
        self.dialog14.setWindowModality(Qt.NonModal)
        self.dialog14.show()

    def extButtonClicked4(self):
        project_id = self.selected_project_id
        temp_N = self.D4_N.text()
        temp_TE = self.D4_TE.text()

        if temp_N == '':
            self.alertbox_open()
            return

        cursor = self.cnxn.cursor()

        sql_query = '''
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
                '''.format(field=project_id)

        self.dataframe = pd.read_sql(sql_query, self.cnxn)

        model = DataFrameModel(self.dataframe)
        self.viewtable.setModel(model)

    def extButtonClicked5(self):
        project_id = self.selected_project_id
        temp_start = self.D5_term_start.text()
        temp_end = self.D5_term_end.text()

        if temp_start == '':
            self.alertbox_open()
            return

        cursor = self.cnxn.cursor()

        sql_query = '''
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
                '''.format(field=project_id)

        self.dataframe = pd.read_sql(sql_query, self.cnxn)

        model = DataFrameModel(self.dataframe)
        self.viewtable.setModel(model)

    def extButtonClicked6(self):
        project_id = self.selected_project_id
        tempDate = self.D6_Date.text()  # 필수값
        tempTDate = self.D6_Date2.text()  # 필수값
        tempAccount = self.D6_Account.text()
        tempJE = self.D6_JE.text()
        tempCost = self.D6_Cost.text()

        if tempDate == '--' or tempTDate == '':
            self.alertbox_open()
            return

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
                '''.format(field=project_id)

        self.dataframe = pd.read_sql(sql, self.cnxn)

        model = DataFrameModel(self.dataframe)
        self.viewtable.setModel(model)

    def extButtonClicked7(self):
        project_id = self.selected_project_id
        tempDate = self.D7_Date.text()  # 필수값
        tempAccount = self.D7_Account.text()
        tempJE = self.D7_JE.text()
        tempCost = self.D7_Cost.text()

        if self.rbtn1.isChecked():
            tempState = 'Effective Date'

        elif self.rbtn2.isChecked():
            tempState = 'Entry Date'

        if tempDate == '--':
            self.alertbox_open()
            return

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
                '''.format(field=project_id)

        self.dataframe = pd.read_sql(sql, self.cnxn)

        model = DataFrameModel(self.dataframe)
        self.viewtable.setModel(model)

    def extButtonClicked8(self):
        project_id = self.selected_project_id
        tempN = self.D8_N.text()  # 필수값
        tempAccount = self.D8_Account.text()
        tempJE = self.D8_JE.text()
        tempCost = self.D8_Cost.text()

        if tempN == '':
            self.alertbox_open()
            return

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
                '''.format(field=project_id)

        self.dataframe = pd.read_sql(sql, self.cnxn)

        model = DataFrameModel(self.dataframe)
        self.viewtable.setModel(model)

    def extButtonClicked9(self):
        project_id = self.selected_project_id
        tempN = self.D9_N.text()  # 필수값
        tempTE = self.D9_TE.text()

        if tempN == '' or tempTE == '':
            self.alertbox_open()
            return

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
            '''.format(field=project_id)

        self.dataframe = pd.read_sql(sql, self.cnxn)

        model = DataFrameModel(self.dataframe)
        self.viewtable.setModel(model)

    def extButtonClicked10(self):
        project_id = self.selected_project_id
        tempSearch = self.D10_Search.text()  # 필수값
        tempAccount = self.D10_Account.text()
        tempPoint = self.D10_Point.text()
        tempTE = self.D10_TE.text()

        if tempSearch == '' or tempAccount == '' or tempPoint == '' or tempTE == '':
            self.alertbox_open()
            return

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
            '''.format(field=project_id)

        self.dataframe = pd.read_sql(sql, self.cnxn)

        model = DataFrameModel(self.dataframe)
        self.viewtable.setModel(model)

    # A, B, C조 추후 논의
    # def extButtonClicked13(self):
    #     return
    # def extButtonClicked13(self):
    #     return
    #
    def extButtonClicked13(self):
        project_id = self.selected_project_id
        temp_continuous = self.D13_continuous.text()  # 필수
        temp_account_13 = self.D13_account.text()
        temp_cost_13 = self.D13_cost.text()

        if temp_continuous == '' or temp_account_13 == '' or temp_cost_13 == '':
            self.alertbox_open()
            return

        cursor = self.cnxn.cursor()

        # sql문 수정
        sql_query = '''
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
            '''.format(field=project_id)

        self.dataframe = pd.read_sql(sql_query, self.cnxn)
        model = DataFrameModel(self.dataframe)
        self.viewtable.setModel(model)

    def extButtonClicked14(self):

        project_id = self.selected_project_id
        tempKey = self.D14_Key.text()  # 필수값
        tempTE = self.D14_TE.text()

        if tempKey == '' or tempTE == '':
            self.alertbox_open()
            return

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
            '''.format(field=project_id)

        self.dataframe = pd.read_sql(sql, self.cnxn)

        model = DataFrameModel(self.dataframe)
        self.viewtable.setModel(model)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())
