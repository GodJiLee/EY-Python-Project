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
        self.users = None
        self.cnxn = None
        self.selected_senario_class = None
        self.selected_senario_subclass = None

    def init_UI(self):

        img = QImage('./dark_gray.png')
        scaledImg = img.scaled(QSize(1000, 700))
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
        if server == "--서버 목록--":
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
            if len(selected_project_names) == 0:
                QMessageBox.about(self, "Warning", "해당 Engagement Code 내 프로젝트가 존재하지 않습니다.")
                self.ProjectCombobox.clear()
                self.ProjectCombobox.addItem("프로젝트가 없습니다.")

            else:
                self.ProjectCombobox.clear()
                self.ProjectCombobox.addItem("--프로젝트 목록--")

                for i in range(0, len(selected_project_names)):
                    self.ProjectCombobox.addItem(selected_project_names.iloc[i, 0])

            return

        except:
            QMessageBox.about(self, "Warning", "Engagement Code를 입력하세요.")
            self.ProjectCombobox.clear()
            self.ProjectCombobox.addItem("프로젝트가 없습니다")
            return

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

    def Project_ID_Selected(self, text):
        self.selected_project_id = text

    def connectDialog(self):
        if self.selected_project_id == '--프로젝트 목록--' or self.selected_project_id == None:
            QMessageBox.about(self, "Warning", "프로젝트를 선택하세요.")
        else:
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
                QMessageBox.about(self, "Warning", "시나리오를 선택하세요.")

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
        self.cb.addItem('--서버 목록--')
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
        btn1.setStyleSheet('color:white;  background-image : url(./bar.png)')

        btn1.clicked.connect(self.connectButtonClicked)

        self.ProjectCombobox = QComboBox(self)
        self.cb.activated[str].connect(self.onActivated)
        self.ProjectCombobox.activated[str].connect(self.projectselected)
        self.le3 = QLineEdit(self)

        btn2 = QPushButton('   Input Conditions', self)
        font2 = btn2.font()
        font2.setBold(True)
        btn2.setFont(font2)
        btn2.setStyleSheet('color:white;  background-image : url(./bar.png)')

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
        self.btn2.clicked.connect(self.extButtonClicked5_No_SAP)

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
        self.btnDialog.setStyleSheet(
            'color:white;  background-image : url(./bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close5)

        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)

        self.btnDialog1 = QPushButton('Close', self.dialog5)
        self.btnDialog1.setStyleSheet(
            'color:white;  background-image : url(./bar.png)')
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


        label_SAP_Example = QLabel('※ SKA1 파일을 Drop하시오', self.dialog5)
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
        self.D6_Date.setInputMask("0000-00-00;*")

        labelDate2 = QLabel('T일* : ', self.dialog6)
        labelDate2.setStyleSheet("color: white;")

        font2 = labelDate2.font()
        font2.setBold(True)
        labelDate2.setFont(font2)

        self.D6_Date2 = QLineEdit(self.dialog6)
        self.D6_Date2.setStyleSheet("background-color: white;")

        labelAccount = QLabel('특정계정 : ', self.dialog6)
        labelAccount.setStyleSheet("color: white;")

        font3 = labelAccount.font()
        font3.setBold(True)
        labelAccount.setFont(font3)

        self.D6_Account = QLineEdit(self.dialog6)
        self.D6_Account.setStyleSheet("background-color: white;")

        labelJE = QLabel('전표입력자 : ', self.dialog6)
        labelJE.setStyleSheet("color: white;")

        font4 = labelJE.font()
        font4.setBold(True)
        labelJE.setFont(font4)

        self.D6_JE = QLineEdit(self.dialog6)
        self.D6_JE.setStyleSheet("background-color: white;")

        labelCost = QLabel('중요성금액 : ', self.dialog6)
        labelCost.setStyleSheet("color: white;")

        font5 = labelCost.font()
        font5.setBold(True)
        labelCost.setFont(font5)

        self.D6_Cost = QLineEdit(self.dialog6)
        self.D6_Cost.setStyleSheet("background-color: white;")

        self.D6_Date.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D6_Date2.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D6_Account.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D6_JE.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D6_Cost.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        layout1 = QGridLayout()
        layout1.addWidget(labelDate, 0, 0)
        layout1.addWidget(self.D6_Date, 0, 1)
        layout1.addWidget(labelDate2, 1, 0)
        layout1.addWidget(self.D6_Date2, 1, 1)
        layout1.addWidget(labelAccount, 2, 0)
        layout1.addWidget(self.D6_Account, 2, 1)
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
        self.dialog6.setGeometry(300, 300, 500, 250)

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
        self.D7_Date.setInputMask("0000-00-00;*")
        self.D7_Date.setStyleSheet("background-color: white;")

        labelAccount = QLabel('특정계정 : ', self.dialog7)
        labelAccount.setStyleSheet("color: white;")

        font4 = labelAccount.font()
        font4.setBold(True)
        labelAccount.setFont(font4)

        self.D7_Account = QLineEdit(self.dialog7)
        self.D7_Account.setStyleSheet("background-color: white;")

        labelJE = QLabel('전표입력자 : ', self.dialog7)
        labelJE.setStyleSheet("color: white;")

        font5 = labelJE.font()
        font5.setBold(True)
        labelJE.setFont(font5)

        self.D7_JE = QLineEdit(self.dialog7)
        self.D7_JE.setStyleSheet("background-color: white;")

        labelCost = QLabel('중요성금액 : ', self.dialog7)
        labelCost.setStyleSheet("color: white;")

        font6 = labelCost.font()
        font6.setBold(True)
        labelCost.setFont(font6)

        self.D7_Cost = QLineEdit(self.dialog7)
        self.D7_Cost.setStyleSheet("background-color: white;")

        self.D7_Date.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D7_Account.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D7_JE.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D7_Cost.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        layout0 = QGridLayout()
        layout0.addWidget(self.rbtn1, 0, 0)
        layout0.addWidget(self.rbtn2, 0, 1)

        layout1 = QGridLayout()
        layout1.addWidget(labelDate, 0, 0)
        layout1.addWidget(self.D7_Date, 0, 1)
        layout1.addWidget(labelAccount, 1, 0)
        layout1.addWidget(self.D7_Account, 1, 1)
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
        self.dialog7.setGeometry(300, 300, 500, 200)
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

        labelAccount = QLabel('특정계정 : ', self.dialog8)
        labelAccount.setStyleSheet("color: white;")

        font2 = labelAccount.font()
        font2.setBold(True)
        labelAccount.setFont(font2)

        self.D8_Account = QLineEdit(self.dialog8)
        self.D8_Account.setStyleSheet("background-color: white;")

        labelJE = QLabel('전표입력자 : ', self.dialog8)
        labelJE.setStyleSheet("color: white;")

        font3 = labelJE.font()
        font3.setBold(True)
        labelJE.setFont(font3)

        self.D8_JE = QLineEdit(self.dialog8)
        self.D8_JE.setStyleSheet("background-color: white;")

        labelCost = QLabel('중요성금액 : ', self.dialog8)
        labelCost.setStyleSheet("color: white;")

        font4 = labelCost.font()
        font4.setBold(True)
        labelCost.setFont(font4)

        self.D8_Cost = QLineEdit(self.dialog8)
        self.D8_Cost.setStyleSheet("background-color: white;")

        self.D8_N.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D8_Account.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D8_JE.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D8_Cost.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        layout1 = QGridLayout()
        layout1.addWidget(labelDate, 0, 0)
        layout1.addWidget(self.D8_N, 0, 1)
        layout1.addWidget(labelAccount, 1, 0)
        layout1.addWidget(self.D8_Account, 1, 1)
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
        self.dialog8.setGeometry(300, 300, 500, 200)

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

    # A,B,C조 작성 - 추후 논의
    # def Dialog12(self):

    def Dialog13(self):
        self.dialog13 = QDialog()
        self.dialog13.setStyleSheet('background-color: #2E2E38')
        self.dialog13.setWindowIcon(QIcon('./EY_logo.png'))

        self.btn2 = QPushButton('   Extract Data', self.dialog13)
        self.btn2.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btn2.clicked.connect(self.extButtonClicked13)

        font9 = self.btn2.font()
        font9.setBold(True)
        self.btn2.setFont(font9)

        self.btnDialog = QPushButton('   Close', self.dialog13)
        self.btnDialog.setStyleSheet(
            'color:white;  background-image : url(./bar.png)')
        self.btnDialog.clicked.connect(self.dialog_close13)

        font10 = self.btnDialog.font()
        font10.setBold(True)
        self.btnDialog.setFont(font10)

        self.btn2.resize(110, 30)
        self.btnDialog.resize(110, 30)

        labelContinuous = QLabel('연속된 자릿수* : ', self.dialog13)
        labelContinuous.setStyleSheet("color: white;")

        font1 = labelContinuous.font()
        font1.setBold(True)
        labelContinuous.setFont(font1)

        self.D13_continuous = QLineEdit(self.dialog13)
        self.D13_continuous.setStyleSheet("background-color: white;")

        labelAccount = QLabel('특정계정 : ', self.dialog13)
        labelAccount.setStyleSheet("color: white;")

        font2 = labelAccount.font()
        font2.setBold(True)
        labelAccount.setFont(font2)

        self.D13_account = QLineEdit(self.dialog13)
        self.D13_account.setStyleSheet("background-color: white;")

        labelCost = QLabel('중요성금액 : ', self.dialog13)
        labelCost.setStyleSheet("color: white;")

        font3 = labelCost.font()
        font3.setBold(True)
        labelCost.setFont(font3)

        self.D13_cost = QLineEdit(self.dialog13)
        self.D13_cost.setStyleSheet("background-color: white;")

        self.D13_continuous.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D13_account.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소
        self.D13_cost.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        layout1 = QGridLayout()
        layout1.addWidget(labelContinuous, 0, 0)
        layout1.addWidget(self.D13_continuous, 0, 1)
        layout1.addWidget(labelAccount, 1, 0)
        layout1.addWidget(self.D13_account, 1, 1)
        layout1.addWidget(labelCost, 2, 0)
        layout1.addWidget(self.D13_cost, 2, 1)

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

        self.dialog13.setLayout(main_layout)
        self.dialog13.setGeometry(300, 300, 500, 170)

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
    def AddSheetButton_Clicked(self):
        return

    def RemoveSheetButton_Clicked(self):
        return

    def Sheet_ComboBox_Selected(self):
        return

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


        ##Add Sheet 버튼
        AddSheet_button = QPushButton('Add Sheet')
        AddSheet_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        AddSheet_button.setStyleSheet('color:white;background-image : url(./bar.png)')
        font_AddSheet = AddSheet_button.font()
        font_AddSheet.setBold(True)
        AddSheet_button.setFont(font_AddSheet)

        ##RemoveSheet 버튼
        RemoveSheet_button = QPushButton('Remove Sheet')
        RemoveSheet_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        RemoveSheet_button.setStyleSheet('color:white;background-image : url(./bar.png)')
        font_RemoveSheet = RemoveSheet_button.font()
        font_RemoveSheet.setBold(True)
        RemoveSheet_button.setFont(font_RemoveSheet)


        #label
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
        AddSheet_button.clicked.connect(self.AddSheetButton_Clicked)
        RemoveSheet_button.clicked.connect(self.RemoveSheetButton_Clicked)
        save_path_button.clicked.connect(self.saveFileDialog)
        export_file_button.clicked.connect(self.saveFile)
        self.combo_sheet.currentIndexChanged.connect(self.Sheet_ComboBox_Selected)


        ##layout 쌓기
        left_sublayout.addWidget(label_sheet, 0, 0)
        left_sublayout.addWidget(self.combo_sheet, 0, 1)
        left_sublayout.addWidget(label_savepath, 1, 0)
        left_sublayout.addWidget(self.line_savepath, 1, 1)

        right_sublayout2.addWidget(AddSheet_button, stretch=1)
        right_sublayout2.addWidget(RemoveSheet_button, stretch=1)
        right_sublayout3.addWidget(save_path_button, stretch=1)
        right_sublayout3.addWidget(export_file_button, stretch=1)

        right_sublayout1.addLayout(right_sublayout2, stretch=1)
        right_sublayout1.addLayout(right_sublayout3, stretch=1)

        layout.addLayout(left_sublayout, stretch=2)
        layout.addLayout(right_sublayout1, stretch=1)

        groupbox.setLayout(layout)

        return groupbox


    def projectselected(self, text):
        self.Project_ID_Selected(text)
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

        self.dataframe = pd.read_sql(sql_query, self.cnxn)

        model = DataFrameModel(self.dataframe)
        self.viewtable.setModel(model)
    
    def extButtonClicked5_SAP(self):
        return
    
    def extButtonClicked5_No_SAP(self):
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

            self.dataframe = pd.read_sql(sql, self.cnxn)

            model = DataFrameModel(self.dataframe)
            self.viewtable.setModel(model)

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

            self.dataframe = pd.read_sql(sql, self.cnxn)

            model = DataFrameModel(self.dataframe)
            self.viewtable.setModel(model)

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

            self.dataframe = pd.read_sql(sql, self.cnxn)

            model = DataFrameModel(self.dataframe)
            self.viewtable.setModel(model)

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

            self.dataframe = pd.read_sql(sql, self.cnxn)

            model = DataFrameModel(self.dataframe)
            self.viewtable.setModel(model)

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

            self.dataframe = pd.read_sql(sql, self.cnxn)

            model = DataFrameModel(self.dataframe)
            self.viewtable.setModel(model)

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

        self.dataframe = pd.read_sql(sql_query, self.cnxn)
        model = DataFrameModel(self.dataframe)
        self.viewtable.setModel(model)

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

            self.dataframe = pd.read_sql(sql, self.cnxn)

            model = DataFrameModel(self.dataframe)
            self.viewtable.setModel(model)

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
            self.line_savepath.setText(saveRoute)

    def saveFile(self):
        if self.dataframe is None:
            QMessageBox.about(self, "Warning", "저장할 데이터가 없습니다")
            return

        else:
            self.dataframe.to_csv(saveRoute, index=False, encoding='utf-8-sig')
            QMessageBox.about(self, "Warning", "데이터를 저장했습니다.")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())

