### Dialog
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
    self.btnDialog.setStyleSheet('color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
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
    self.dialog5.setStyleSheet('background-color: #2E2E38')
    self.dialog5.setWindowIcon(QIcon('C:/Users/BZ297TR/OneDrive - EY/Desktop/EY_logo.png'))

    self.btn2 = QPushButton(' Extract Data', self.dialog5)
    self.btn2.move(70, 200)
    self.btn2.setStyleSheet('color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
    self.btn2.clicked.connect(self.extButtonClicked5)

    font9 = self.btn2.font()
    font9.setBold(True)
    self.btn2.setFont(font9)

    self.btnDialog = QPushButton('Close', self.dialog5)
    self.btnDialog.move(180, 200)
    self.btnDialog.setStyleSheet('color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
    self.btnDialog.clicked.connect(self.dialog_close5)

    font10 = self.btnDialog.font()
    font10.setBold(True)
    self.btnDialog.setFont(font10)

    label_term_start = QLabel('시작일 :', self.dialog5)
    label_term_start.move(50, 80)
    label_term_start.setStyleSheet('color: white;')

    font1 = label_term_start.font()
    font1.setBold(True)
    label_term_start.setFont(font1)

    self.D5_term_start = QLineEdit(self.dialog5)
    self.D5_term_start.setInputMask('0000-00-00;*')
    self.D5_term_start.setStyleSheet('background-color: white;')
    self.D5_term_start.move(150, 80)

    label_term_end = QLabel('종료일 :', self.dialog5)
    label_term_end.move(50, 110)
    label_term_end.setStyleSheet('color: white;')

    font2 = label_term_end.font()
    font2.setBold(True)
    label_term_end.setFont(font2)

    self.D5_term_end = QLineEdit(self.dialog5)
    self.D5_term_end.setInputMask('0000-00-00;*')
    self.D5_term_end.setStyleSheet('background-color: white;')
    self.D5_term_end.move(150, 110)

    self.dialog5.setGeometry(300, 300, 350, 300)

    self.dialog5.setWindowTitle('Scenario5')
    self.dialog5.setWindowModality(Qt.NonModal)
    self.dialog5.show()


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
    self.btnDialog.setStyleSheet('color:white;  background-image : url(C:/Users/BZ297TR/OneDrive - EY/Desktop/bar.png)')
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

#### extButtonClicked
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

    temp_start = self.D5_term_start.text()
    temp_end = self.D5_term_end.text()

    if temp_start == '':
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

    df = pd.read_sql(sql_query, cnxn)

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
