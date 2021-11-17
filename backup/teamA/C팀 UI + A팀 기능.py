import sys

from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import pyodbc
import pandas as pd


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
        if not index.isValid() or not (0 <= index.row() < self.rowCount() \
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
        self.init_UI()
        ##Initialize Variables
        self.selected_project_id = None
        self.selected_server_name = "서버를 선택하세요"
        self.dataframe = None
        self.users = None
        self.cnxn = None
        self.selected_senario_class = None
        self.selected_senario_subclass = None

    def init_UI(self):

        img = QImage("C:/Users/BZ297TR/OneDrive - EY/Desktop//dark_gray.png")
        scaledImg = img.scaled(QSize(1000, 700))
        palette = QPalette()
        palette.setBrush(10, QBrush(scaledImg))
        self.setPalette(palette)

        pixmap = QPixmap('C:/Users/BZ297TR/OneDrive - EY/Desktop/title.png')
        lbl_img = QLabel()
        lbl_img.setPixmap(pixmap)

        widget_layout = QBoxLayout(QBoxLayout.TopToBottom)
        self.splitter_layout = QSplitter(Qt.Vertical)

        self.splitter_layout.addWidget(lbl_img)
        self.splitter_layout.addWidget(self.Connect_ServerInfo_Group())
        self.splitter_layout.addWidget(self.Show_DataFrame_Group())
        self.splitter_layout.addWidget(self.Save_Buttons_Group())

        widget_layout.addWidget(self.splitter_layout)
        self.setLayout(widget_layout)

        self.setWindowIcon(QIcon("C:/Users/BZ297TR/OneDrive - EY/Desktop/EY_logo.png"))
        self.setWindowTitle('Scenario')

        self.setGeometry(300, 300, 1000, 700)
        self.show()

    def Server_ComboBox_Selected(self, text):
        self.selected_server_name = text

    def connectButtonClicked(self):
        global passwords
        global users

        passwords = ''
        ecode = self.le3.text()
        users = 'guest'
        server = self.selected_server_name
        password = passwords
        db = 'master'
        user = users

        # 예외처리 - 서버 선택
        if server == "서버를 선택하세요":
            QMessageBox.about(self, "Warning", "서버가 선택되어 있지 않습니다")
            return

        # 예외처리 - 접속 정보 오류
        try:
            self.cnxn = pyodbc.connect("DRIVER={SQL Server};SERVER=" +
                                       server + ";uid=" + user + ";pwd=" +
                                       password + ";DATABASE=" + db +
                                       ";trusted_connection=" + "yes")
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

        # 예외처리 - ecode, pname 오류
        try:
            selected_project_names = pd.read_sql(sql_query, self.cnxn)
        except:
            QMessageBox.about(self, "Warning", "Engagement Code가 잘못되었습니다")
            self.ProjectCombobox.clear()
            self.ProjectCombobox.addItem("프로젝트가 없습니다")
            return

        # 예외처리 - pname 입력 오류
        if len(selected_project_names) == 0:
            QMessageBox.about(self, "Warning", "프로젝트가 없습니다")
            self.ProjectCombobox.clear()
            self.ProjectCombobox.addItem("프로젝트가 없습니다")
            return

        self.ProjectCombobox.clear()
        self.ProjectCombobox.addItem("--프로젝트 목록--")

        for i in range(0, len(selected_project_names)):
            self.ProjectCombobox.addItem(selected_project_names.iloc[i, 0])

    def onActivated(self, text):
        global ids
        ids = text

    def updateSmallCombo(self, index):
        global big
        big = index
        self.comboSmall.clear()
        sc = self.comboBig.itemData(index)
        if sc:
            self.comboSmall.addItems(sc)

    def saveSmallCombo(self, index):
        global small
        small = index

    def connectDialog(self):
        if big == 0 and small == 1:
            self.Dialog4()

        elif big == 0 and small == 2:
            self.Dialog5()

        elif big == 1 and small == 1:
            self.Dialog6()

        elif big == 1 and small == 2:
            self.Dialog7()

        elif big == 1 and small == 3:
            self.Dialog8()

        elif big == 2 and small == 1:
            self.Dialog9()

        elif big == 2 and small == 2:
            self.Dialog10()

        elif big == 3 and small == 1:
            self.Dialog11()

        elif big == 3 and small == 2:
            self.Dialog12()

        elif big == 4 and small == 1:
            self.Dialog13()

        elif big == 4 and small == 2:
            self.Dialog14()

        else:
            self.messagebox_open()

    def Connect_ServerInfo_Group(self):
        groupbox = QGroupBox('접속 정보')
        self.setStyleSheet('QGroupBox  {color: white;}')
        font5 = groupbox.font()
        font5.setBold(True)
        groupbox.setFont(font5)

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

        self.cb = QComboBox(self)
        self.cb.addItem('KRSEOVMPPACSQ01\INST1')
        self.cb.addItem('KRSEOVMPPACSQ02\INST1')
        self.cb.addItem('KRSEOVMPPACSQ03\INST1')
        self.cb.addItem('KRSEOVMPPACSQ04\INST1')
        self.cb.addItem('KRSEOVMPPACSQ05\INST1')
        self.cb.addItem('KRSEOVMPPACSQ06\INST1')
        self.cb.addItem('KRSEOVMPPACSQ07\INST1')
        self.cb.addItem('KRSEOVMPPACSQ08\INST1')
        self.cb.move(50, 50)

        self.cb.activated[str].connect(self.Server_ComboBox_Selected)

        self.comboBig = QComboBox(self)

        self.comboBig.addItem('Data 완전성', ['--시나리오 목록--', '계정 사용빈도 N번이하인 계정이 포함된 전표리스트', '당기 생성된 계정리스트 추출'])
        self.comboBig.addItem('Data Timing',
                              ['--시나리오 목록--', '결산일 전후 T일 입력 전표', '영업일 전기/입력 전표', '효력, 입력 일자 간 차이가 N일 이상인 전표'])
        self.comboBig.addItem('Data 업무분장',
                              ['--시나리오 목록--', '전표 작성 빈도수가 N회 이하인 작성자에 의한 생성된 전표', '특정 전표 입력자(W)에 의해 생성된 전표'])
        self.comboBig.addItem('Data 분개검토',
                              ['--시나리오 목록--', '특정한 주계정(A)과 특정한 상대계정(B)이 아닌 전표리스트 검토', '특정 계정(A)이 감소할 때 상대계정 리스트 검토'])
        self.comboBig.addItem('기타', ['--시나리오 목록--', '연속된 숫자로 끝나는 금액 검토',
                                     '전표 description에 공란 또는 특정단어(key word)가 입력되어 있는 전표 리스트 (TE금액 제시 가능)'])

        self.comboSmall = QComboBox(self)

        self.comboBig.currentIndexChanged.connect(self.updateSmallCombo)
        self.updateSmallCombo(self.comboBig.currentIndex())

        self.comboSmall.currentIndexChanged.connect(self.saveSmallCombo)
        self.saveSmallCombo(self.comboSmall.currentIndex())

        btn1 = QPushButton('   SQL Server Connect', self)
        font1 = btn1.font()
        font1.setBold(True)
        btn1.setFont(font1)
        btn1.setStyleSheet('color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')

        btn1.clicked.connect(self.connectButtonClicked)

        self.ProjectCombobox = QComboBox(self)
        self.cb.activated[str].connect(self.onActivated)
        self.ProjectCombobox.activated[str].connect(self.projectselected)
        self.le3 = QLineEdit(self)

        btn2 = QPushButton('   Input Conditions', self)
        font2 = btn2.font()
        font2.setBold(True)
        btn2.setFont(font2)
        btn2.setStyleSheet('color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')

        btn2.clicked.connect(self.connectDialog)

        grid = QGridLayout()
        grid.addWidget(label1, 0, 0)
        grid.addWidget(label2, 1, 0)
        grid.addWidget(label3, 2, 0)
        grid.addWidget(label4, 3, 0)
        grid.addWidget(self.cb, 0, 1)
        grid.addWidget(btn1, 0, 2)
        grid.addWidget(self.comboBig, 3, 1)
        grid.addWidget(self.comboSmall, 4, 1)
        grid.addWidget(btn2, 4, 2)
        grid.addWidget(self.le3, 1, 1)
        grid.addWidget(self.ProjectCombobox, 2, 1)

        groupbox.setLayout(grid)

        return groupbox

    def Dialog4(self):
        self.dialog4 = QDialog()
        self.dialog4.setStyleSheet('background-color: #2E2E38')
        self.dialog4.setWindowIcon(QIcon('C:/Users/BZ297TR/OneDrive - EY/Desktop/EY_logo.png'))

        self.btn2 = QPushButton(' Extract Data', self.dialog4)
        self.btn2.move(70, 200)
        self.btn2.setStyleSheet('color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked4)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        self.btnDialog = QPushButton('Close', self.dialog4)
        self.btnDialog.move(180, 200)
        self.btnDialog.setStyleSheet(
            'color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close4)

        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)

        label_freq = QLabel('사용빈도(N)* :', self.dialog4)
        label_freq.move(50, 80)
        label_freq.setStyleSheet('color: white;')

        font1 = label_freq.font()
        font1.setBold(True)
        label_freq.setFont(font1)

        self.D4_N = QLineEdit(self.dialog4)
        self.D4_N.setStyleSheet('background-color: white;')
        self.D4_N.move(150, 80)

        label_TE = QLabel('중요성금액: ', self.dialog4)
        label_TE.move(50, 110)
        label_TE.setStyleSheet('color: white;')

        font2 = label_TE.font()
        font2.setBold(True)
        label_TE.setFont(font2)

        self.D4_TE = QLineEdit(self.dialog4)
        self.D4_TE.setStyleSheet('background-color: white;')
        self.D4_TE.move(150, 110)

        self.dialog4.setGeometry(300, 300, 350, 300)

        self.dialog4.setWindowTitle('Scenario4')
        self.dialog4.setWindowModality(Qt.NonModal)
        self.dialog4.show()

    def Dialog5(self):
        self.dialog5 = QDialog()
        self.dialog5.resize(200, 200)
        self.dialog5.setStyleSheet('background-color: #2E2E38')
        self.dialog5.setWindowIcon(QIcon('C:/Users/BZ297TR/OneDrive - EY/Desktop/EY_logo.png'))

        ### Extract Data
        self.btn2 = QPushButton(' Extract Data', self.dialog5)
        self.btn2.move(90, 330)
        self.btn2.setStyleSheet('color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked5)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        ### Close
        self.btnDialog = QPushButton('Close', self.dialog5)
        self.btnDialog.move(250, 330)
        self.btnDialog.setStyleSheet(
            'color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close5)

        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)

        self.btn2.resize(110, 30)
        self.btnDialog.resize(110, 30)

        ### 라벨1 - 계정코드 입력
        label_AccCode = QLabel('Enter your Account Code: ', self.dialog5)
        label_AccCode.move(20, 40)
        label_AccCode.setStyleSheet('color: white;')
        label_AccCode.setFont(QFont('Arial', 12))

        font1 = label_AccCode.font()
        font1.setBold(True)
        label_AccCode.setFont(font1)

        ### 라벨 2 - 입력 예시
        label_Example = QLabel('※ 입력 예시: OO', self.dialog5)
        label_Example.move(40, 70)
        label_Example.setStyleSheet('color: red;')
        label_Example.setFont(QFont('Times font', 9))

        font2 = label_Example.font()
        font2.setBold(False)
        label_Example.setFont(font2)

        ### TextEdit - 계정코드 Paste
        self.D5_Code = QTextEdit(self.dialog5)
        self.D5_Code.setAcceptRichText(False)
        self.D5_Code.resize(350, 200)
        self.D5_Code.move(60, 110)
        self.D5_Code.setStyleSheet('color: black;')
        self.D5_Code.setStyleSheet('background-color: white;')

        self.dialog5.setGeometry(400, 400, 465, 400)

        self.dialog5.setWindowTitle('Scenario5')
        self.dialog5.setWindowModality(Qt.NonModal)
        self.dialog5.show()

    def Dialog6(self):
        self.dialog6 = QDialog()
        self.dialog6.setStyleSheet('background-color: #2E2E38')
        self.dialog6.setWindowIcon(QIcon("C:/Users/BZ297TR/OneDrive - EY/Desktop/EY_logo.png"))

        ### Extract Data
        self.btn2 = QPushButton(' Extract Data', self.dialog6)
        self.btn2.move(70, 200)
        self.btn2.setStyleSheet('color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked6)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        self.btnDialog = QPushButton(" Close", self.dialog6)
        self.btnDialog.move(180, 200)
        self.btnDialog.setStyleSheet('color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
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

    def Dialog7(self):
        self.dialog7 = QDialog()
        self.dialog7.setStyleSheet('background-color: #2E2E38')
        self.dialog7.setWindowIcon(QIcon("C:/Users/BZ297TR/OneDrive - EY/Desktop/EY_logo.png"))

        self.btn2 = QPushButton(' Extract Data', self.dialog7)
        self.btn2.move(80, 200)
        self.btn2.setStyleSheet('color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked7)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        self.btnDialog = QPushButton(" Close", self.dialog7)
        self.btnDialog.move(190, 200)
        self.btnDialog.setStyleSheet('color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
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

    def Dialog8(self):
        self.dialog8 = QDialog()
        self.dialog8.setStyleSheet('background-color: #2E2E38')
        self.dialog8.setWindowIcon(QIcon("C:/Users/BZ297TR/OneDrive - EY/Desktop/EY_logo.png"))

        self.btn2 = QPushButton(' Extract Data', self.dialog8)
        self.btn2.move(60, 180)
        self.btn2.setStyleSheet('color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked8)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        self.btnDialog = QPushButton(" Close", self.dialog8)
        self.btnDialog.move(170, 180)
        self.btnDialog.setStyleSheet('color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
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

    def Dialog9(self):
        self.dialog9 = QDialog()
        self.dialog9.setStyleSheet('background-color: #2E2E38')
        self.dialog9.setWindowIcon(QIcon("C:/Users/BZ297TR/OneDrive - EY/Desktop/EY_logo.png"))

        self.btn2 = QPushButton(' Extract Data', self.dialog9)
        self.btn2.move(60, 180)
        self.btn2.setStyleSheet('color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked9)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        self.btnDialog = QPushButton(" Close", self.dialog9)
        self.btnDialog.move(170, 180)
        self.btnDialog.setStyleSheet('color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close9)

        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)

        labelKeyword = QLabel('작성빈도수 : ', self.dialog9)
        labelKeyword.move(50, 50)
        labelKeyword.setStyleSheet("color: white;")

        font1 = labelKeyword.font()
        font1.setBold(True)
        labelKeyword.setFont(font1)

        self.D9_N = QLineEdit(self.dialog9)
        self.D9_N.setStyleSheet("background-color: white;")
        self.D9_N.move(150, 50)

        labelTE = QLabel('TE : ', self.dialog9)
        labelTE.move(50, 80)
        labelTE.setStyleSheet("color: white;")

        font2 = labelTE.font()
        font2.setBold(True)
        labelTE.setFont(font2)

        self.D9_TE = QLineEdit(self.dialog9)
        self.D9_TE.setStyleSheet("background-color: white;")
        self.D9_TE.move(150, 80)

        self.dialog9.setWindowTitle("Scenario9")
        self.dialog9.setWindowModality(Qt.NonModal)
        self.dialog9.show()

    def Dialog10(self):
        self.dialog10 = QDialog()
        self.dialog10.setStyleSheet('background-color: #2E2E38')
        self.dialog10.setWindowIcon(QIcon("C:/Users/BZ297TR/OneDrive - EY/Desktop/EY_logo.png"))

        self.btn2 = QPushButton(' Extract Data', self.dialog10)
        self.btn2.move(60, 180)
        self.btn2.setStyleSheet('color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked10)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        self.btnDialog = QPushButton(" Close", self.dialog10)
        self.btnDialog.move(170, 180)
        self.btnDialog.setStyleSheet('color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close10)

        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)

        labelKeyword = QLabel('전표입력자 : ', self.dialog10)
        labelKeyword.move(50, 30)
        labelKeyword.setStyleSheet("color: white;")

        font1 = labelKeyword.font()
        font1.setBold(True)
        labelKeyword.setFont(font1)

        self.D10_Search = QLineEdit(self.dialog10)
        self.D10_Search.setStyleSheet("background-color: white;")
        self.D10_Search.move(150, 30)

        labelPoint = QLabel('특정시점 : ', self.dialog10)
        labelPoint.move(50, 60)
        labelPoint.setStyleSheet("color: white;")

        font2 = labelPoint.font()
        font2.setBold(True)
        labelPoint.setFont(font2)

        self.D10_Point = QLineEdit(self.dialog10)
        self.D10_Point.setStyleSheet("background-color: white;")
        self.D10_Point.move(150, 60)

        labelAccount = QLabel('특정계정 : ', self.dialog10)
        labelAccount.move(50, 90)
        labelAccount.setStyleSheet("color: white;")

        font3 = labelAccount.font()
        font3.setBold(True)
        labelAccount.setFont(font3)

        self.D10_Account = QLineEdit(self.dialog10)
        self.D10_Account.setStyleSheet("background-color: white;")
        self.D10_Account.move(150, 90)

        labelTE = QLabel('TE : ', self.dialog10)
        labelTE.move(50, 120)
        labelTE.setStyleSheet("color: white;")

        font4 = labelTE.font()
        font4.setBold(True)
        labelTE.setFont(font4)

        self.D10_TE = QLineEdit(self.dialog10)
        self.D10_TE.setStyleSheet("background-color: white;")
        self.D10_TE.move(150, 120)

        self.dialog10.setWindowTitle("Scenario10")
        self.dialog10.setWindowModality(Qt.NonModal)
        self.dialog10.show()

    # A,B,C조 작성 - 추후 논의
    # def Dialog11(self):

    # A,B,C조 작성 - 추후 논의
    # def Dialog12(self):

    def Dialog13(self):
        self.dialog13 = QDialog()
        self.dialog13.setStyleSheet('background-color: #2E2E38')
        self.dialog13.setWindowIcon(QIcon('C:/Users/BZ297TR/OneDrive - EY/Desktop/EY_logo.png'))

        self.btn2 = QPushButton(' Extract Data', self.dialog13)
        self.btn2.move(70, 200)
        self.btn2.setStyleSheet('color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked13)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        self.btnDialog = QPushButton('Close', self.dialog13)
        self.btnDialog.move(180, 200)
        self.btnDialog.setStyleSheet(
            'color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close13)

        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)

        labelContinuous = QLabel('연속된 자릿수* : ', self.dialog13)
        labelContinuous.move(40, 80)
        labelContinuous.setStyleSheet("color: white;")

        font1 = labelContinuous.font()
        font1.setBold(True)
        labelContinuous.setFont(font1)

        self.D13_continuous = QLineEdit(self.dialog13)
        self.D13_continuous.setStyleSheet("background-color: white;")
        self.D13_continuous.move(170, 80)

        labelAccount = QLabel('특정계정 : ', self.dialog13)
        labelAccount.move(40, 110)
        labelAccount.setStyleSheet("color: white;")

        font2 = labelAccount.font()
        font2.setBold(True)
        labelAccount.setFont(font2)

        self.D13_account = QLineEdit(self.dialog13)
        self.D13_account.setStyleSheet("background-color: white;")
        self.D13_account.move(170, 110)

        labelCost = QLabel('중요성금액 : ', self.dialog13)
        labelCost.move(40, 140)
        labelCost.setStyleSheet("color: white;")

        font3 = labelCost.font()
        font3.setBold(True)
        labelCost.setFont(font3)

        self.D13_cost = QLineEdit(self.dialog13)
        self.D13_cost.setStyleSheet("background-color: white;")
        self.D13_cost.move(170, 140)

        self.dialog13.setGeometry(300, 300, 350, 300)

        self.dialog13.setWindowTitle('Scenario13')
        self.dialog13.setWindowModality(Qt.NonModal)
        self.dialog13.show()

    def Dialog14(self):
        self.dialog14 = QDialog()
        self.dialog14.setStyleSheet('background-color: #2E2E38')
        self.dialog14.setWindowIcon(QIcon("C:/Users/BZ297TR/OneDrive - EY/Desktop/EY_logo.png"))

        self.btn2 = QPushButton(' Extract Data', self.dialog14)
        self.btn2.move(60, 180)
        self.btn2.setStyleSheet('color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked14)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        self.btnDialog = QPushButton(" Close", self.dialog14)
        self.btnDialog.move(170, 180)
        self.btnDialog.setStyleSheet('color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close14)

        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)

        labelKeyword = QLabel('Key Words : ', self.dialog14)
        labelKeyword.move(50, 50)
        labelKeyword.setStyleSheet("color: white;")

        font1 = labelKeyword.font()
        font1.setBold(True)
        labelKeyword.setFont(font1)

        self.D14_Key = QLineEdit(self.dialog14)
        self.D14_Key.setStyleSheet("background-color: white;")
        self.D14_Key.move(150, 50)

        labelTE = QLabel('TE : ', self.dialog14)
        labelTE.move(50, 80)
        labelTE.setStyleSheet("color: white;")

        font2 = labelTE.font()
        font2.setBold(True)
        labelTE.setFont(font2)

        self.D14_TE = QLineEdit(self.dialog14)
        self.D14_TE.setStyleSheet("background-color: white;")
        self.D14_TE.move(150, 80)

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

    def messagebox_open(self):
        self.msg = QMessageBox()
        self.msg.setIcon(QMessageBox.Information)
        self.msg.setWindowTitle('시나리오 재선택')
        self.msg.setText('시나리오를 선택하세요.')
        self.msg.exec_()

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

    def Save_Buttons_Group(self):
        groupbox3 = QGroupBox('저장')
        self.setStyleSheet('QGroupBox  {color: white;}')
        font7 = groupbox3.font()
        font7.setBold(True)
        groupbox3.setFont(font7)

        self.btn3 = QPushButton('Setting Save Route', self)
        self.btn4 = QPushButton('Save', self)
        grid3 = QGridLayout()

        self.btn3.setStyleSheet('color:white;background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
        font1 = self.btn3.font()
        font1.setBold(True)
        self.btn3.setFont(font1)

        self.btn4.setStyleSheet('color:white;background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
        font2 = self.btn4.font()
        font2.setBold(True)
        self.btn4.setFont(font2)

        grid3.addWidget(self.btn3, 0, 0)
        grid3.addWidget(self.btn4, 0, 1)
        self.btn3.clicked.connect(self.saveFileDialog)
        self.btn4.clicked.connect(self.saveFile)

        groupbox3.setLayout(grid3)

        return groupbox3

    def projectselected(self, text):
        passwords = ''
        users = 'guest'

        server = ids
        password = passwords
        db = 'master'
        user = users
        cnxn = pyodbc.connect(
            "DRIVER={SQL Server};SERVER=" + server + ";uid=" + user + ";pwd=" + password + ";DATABASE=" + db + ";trusted_connection=" + "yes")
        cursor = cnxn.cursor()

        sql = '''
                    SELECT Project_ID
                    FROM [DataAnalyticsRepository].[dbo].[Projects]
                    WHERE ProjectName IN (\'{pname}\')
                    AND DeletedBy is Null

                '''.format(pname=text)

        fieldID = pd.read_sql(sql, cnxn)

        global fields
        fields = fieldID.iloc[0, 0]

    def alertbox_open(self):
        self.alt = QMessageBox()
        self.alt.setIcon(QMessageBox.Information)
        self.alt.setWindowTitle('필수 입력값 누락')
        self.alt.setText('필수 입력값이 누락되었습니다.')
        self.alt.exec_()

    def extButtonClicked4(self):
        password = ''
        users = 'guest'
        server = ids  # 서버명
        password = passwords

        temp_N = self.D4_N.text()
        temp_TE = self.D4_TE.text()

        if temp_N == '':
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

        temp_Code = self.D5_Code.text()

        if temp_Code == '':
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

            sql_query = """""".format(field=fields)

        self.dataframe = pd.read_sql(sql_query, self.cnxn)

        model = DataFrameModel(self.dataframe)
        self.viewtable.setModel(model)

    def extButtonClicked6(self):
        passwords = ''
        users = 'guest'
        server = ids
        password = passwords

        tempDate = self.D6_Date.text()  # 필수값
        tempTDate = self.D6_Date2.text()  # 필수값
        tempAccount = self.D6_Account.text()
        tempJE = self.D6_JE.text()
        tempCost = self.D6_Cost.text()

        if tempDate == '--' or tempTDate == '':
            self.alertbox_open()

        else:
            db = 'master'
            user = users
            cnxn = pyodbc.connect(
                "DRIVER={SQL Server};SERVER=" + server + ";uid=" + user + ";pwd=" + password + ";DATABASE=" + db + ";trusted_connection=" + "yes")
            cursor = cnxn.cursor()

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

                '''.format(field=fields)

            df = pd.read_sql(sql, cnxn)

            model = DataFrameModel(df)
            self.viewtable.setModel(model)

            global save_df
            save_df = df

    def extButtonClicked7(self):
        passwords = ''
        users = 'guest'
        server = ids
        password = passwords

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

        else:
            db = 'master'
            user = users
            cnxn = pyodbc.connect(
                "DRIVER={SQL Server};SERVER=" + server + ";uid=" + user + ";pwd=" + password + ";DATABASE=" + db + ";trusted_connection=" + "yes")
            cursor = cnxn.cursor()

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

                '''.format(field=fields)

            df = pd.read_sql(sql, cnxn)

            model = DataFrameModel(df)
            self.viewtable.setModel(model)

            global save_df
            save_df = df

    def extButtonClicked8(self):
        passwords = ''
        users = 'guest'
        server = ids
        password = passwords

        tempN = self.D8_N.text()  # 필수값
        tempAccount = self.D8_Account.text()
        tempJE = self.D8_JE.text()
        tempCost = self.D8_Cost.text()

        if tempN == '':
            self.alertbox_open()

        else:
            db = 'master'
            user = users
            cnxn = pyodbc.connect(
                "DRIVER={SQL Server};SERVER=" + server + ";uid=" + user + ";pwd=" + password + ";DATABASE=" + db + ";trusted_connection=" + "yes")
            cursor = cnxn.cursor()

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

                '''.format(field=fields)

            df = pd.read_sql(sql, cnxn)

            model = DataFrameModel(df)
            self.viewtable.setModel(model)

            global save_df
            save_df = df

    def extButtonClicked9(self):
        passwords = ''
        users = 'guest'
        server = ids
        password = passwords

        tempN = self.D9_N.text()  # 필수값
        tempTE = self.D9_TE.text()

        if tempN == '' or tempTE == '':
            self.alertbox_open()

        else:
            db = 'master'
            user = users
            cnxn = pyodbc.connect(
                "DRIVER={SQL Server};SERVER=" + server + ";uid=" + user + ";pwd=" + password + ";DATABASE=" + db + ";trusted_connection=" + "yes")
            cursor = cnxn.cursor()

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

                '''.format(field=fields)

            df = pd.read_sql(sql, cnxn)

            model = DataFrameModel(df)
            self.viewtable.setModel(model)

            global save_df
            save_df = df

    def extButtonClicked10(self):
        passwords = ''
        users = 'guest'
        server = ids
        password = passwords

        tempSearch = self.D10_Search.text()  # 필수값
        tempAccount = self.D10_Account.text()
        tempPoint = self.D10_Point.text()
        tempTE = self.D10_TE.text()

        if tempSearch == '' or tempAccount == '' or tempPoint == '' or tempTE == '':
            self.alertbox_open()

        else:
            db = 'master'
            user = users
            cnxn = pyodbc.connect(
                "DRIVER={SQL Server};SERVER=" + server + ";uid=" + user + ";pwd=" + password + ";DATABASE=" + db + ";trusted_connection=" + "yes")
            cursor = cnxn.cursor()

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

                '''.format(field=fields)

            df = pd.read_sql(sql, cnxn)

            model = DataFrameModel(df)
            self.viewtable.setModel(model)

            global save_df
            save_df = df

    def extButtonClicked13(self):
        passwords = ''
        users = 'guest'
        server = ids
        password = passwords

        temp_continuous = self.D13_continuous.text()  # 필수
        temp_account_13 = self.D13_account.text()
        temp_cost_13 = self.D13_cost.text()

        if temp_continuous == '' or temp_account_13 == '' or temp_cost_13 == '':
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

        df = pd.read_sql(sql_query, cnxn)
        model = DataFrameModel(df)
        self.viewtable.setModel(model)

        global save_df
        save_df = df

    def extButtonClicked14(self):
        passwords = ''
        users = 'guest'
        server = ids
        password = passwords

        tempKey = self.D14_Key.text()  # 필수값
        tempTE = self.D14_TE.text()

        if tempKey == '' or tempTE == '':
            self.alertbox_open()

        else:
            db = 'master'
            user = users
            cnxn = pyodbc.connect(
                "DRIVER={SQL Server};SERVER=" + server + ";uid=" + user + ";pwd=" + password + ";DATABASE=" + db + ";trusted_connection=" + "yes")
            cursor = cnxn.cursor()

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

                '''.format(field=fields)

            df = pd.read_sql(sql, cnxn)

            model = DataFrameModel(df)
            self.viewtable.setModel(model)

            global save_df
            save_df = df

    @pyqtSlot(QModelIndex)
    def slot_clicked_item(self, QModelIndex):
        self.stk_w.setCurrentIndex(QModelIndex.row())

    def saveFileDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(self, "QFileDialog.getSaveFileName()", "",
                                                  "All Files (*);;Text Files (*.csv)", options=options)
        if fileName:
            global saveRoute
            saveRoute = fileName + ".csv"

    def saveFile(self):
        if save_df is None:
            QMessageBox.about(self, "Warning", "저장할 데이터가 없습니다")
            return

        else:
            save_df.to_csv(saveRoute, index=False, encoding='utf-8-sig')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())
