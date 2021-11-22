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
                    self.new_tree.grandgrandchild.setCheckState(0, Qt.Unchecked)

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
    
    ### 버튼 5 - Save (JE)
    self.btnSave = QPushButton('Save', self.dialog5)
    self.btnSave.setStyleSheet('color: white; background-image: url(./bar.png)')
    self.btnSave.clicked.connect(self.dialog_close5)
    
    font20 = self.btnSave.font()
    font20.setBold(True)
    self.btnSave.setFont(font20)

    self.btnSave.resize(110, 30)
    self.btnSave.resize(110, 30)
    
    ### 버튼 6 - Save and Proceed
    self.btnSaveProceed = QPushButton('Save', self.dialog5)
    self.btnSaveProceed.setStyleSheet('color: white; background-image: url(./bar.png)')
    self.btnSaveProceed.clicked.connect(self.dialog_close5)

    font21 = self.btnSave.font()
    font21.setBold(True)
    self.btnSaveProceed.setFont(font21)

    self.btnSaveProceed.resize(110, 30)
    self.btnSaveProceed.resize(110, 30)

    ### 라벨1 - 계정코드 입력
    label_AccCode = QLabel('Enter your Account Code: ', self.dialog5)
    label_AccCode.setStyleSheet('color: white;')

    font1 = label_AccCode.font()
    font1.setBold(True)
    label_AccCode.setFont(font1)

    # ### 라벨 2 - SKA1 파일 드롭하기
    label_SKA1 = QLabel('※ SKA1 파일을 Drop 하십시오', self.dialog5)
    label_SKA1.setStyleSheet('color: red;')

    font12 = label_SKA1.font()
    font12.setBold(False)
    label_SKA1.setFont(font12)

    ### TextEdit - 계정코드 Paste
    self.MyInput = QTextEdit(self.dialog5)
    self.MyInput.setAcceptRichText(False)
    self.MyInput.setStyleSheet('background-color: white;')
    self.MyInput.setPlaceholderText('※ 입력 예시 : OO')

    ### ListBox Widget
    self.listbox_drops = ListBoxWidget()
    self.listbox_drops.setStyleSheet('background-color: white;')

    ### SAP

    ### 라벨 3 - JE Line Number
    labelJE_Line = QLabel('JE Line Number : ', self.dialog5)
    labelJE_Line.setStyleSheet("color: white;")
    font6 = labelJE_Line.font()
    font6.setBold(True)
    labelJE_Line.setFont(font6)

    ### LineEdit 0 - JE Line Number
    self.D5_JE_Line = QLineEdit(self.dialog5)
    self.D5_JE_Line.setStyleSheet("background-color: white;")
    self.D5_JE_Line.setPlaceholderText('JE Line Number을 입력하세요')

    ### 라벨 4 - JE Number
    labelJE_Number = QLabel('JE Number : ', self.dialog5)
    labelJE_Number.setStyleSheet("color: white;")
    font7 = labelJE_Number.font()
    font7.setBold(True)
    labelJE_Number.setFont(font7)

    ### LineEdit 1 - JE Number
    self.D5_JE_Number = QLineEdit(self.dialog5)
    self.D5_JE_Number.setStyleSheet("background-color: white;")
    self.D5_JE_Number.setPlaceholderText('JE Number를 입력하세요')

    ### 라벨 5 - 시트명 (SAP)
    labelSheet = QLabel('시트명* : ', self.dialog5)
    labelSheet.setStyleSheet("color: white;")
    font5 = labelSheet.font()
    font5.setBold(True)
    labelSheet.setFont(font5)

    ### LineEdit 2 - 시트명 (SAP)
    self.D5_Sheet = QLineEdit(self.dialog5)
    self.D5_Sheet.setStyleSheet("background-color: white;")
    self.D5_Sheet.setPlaceholderText('시트명을 입력하세요')

    ### Non-SAP
    ### 라벨 6 - 시트명 (Non SAP)
    labelSheet2 = QLabel('시트명* : ', self.dialog5)
    labelSheet2.setStyleSheet("color: white;")
    font5 = labelSheet2.font()
    font5.setBold(True)
    labelSheet2.setFont(font5)

    ### LineEdit 3 - 시트명 (Non SAP)
    self.D5_Sheet2 = QLineEdit(self.dialog5)
    self.D5_Sheet2.setStyleSheet("background-color: white;")
    self.D5_Sheet2.setPlaceholderText('시트명을 입력하세요')

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

    ### 탭
    tabs = QTabWidget()
    tab1 = QWidget()
    tab2 = QWidget()
    tab3 = QWidget()

    tab1.setLayout(layout1)
    tab2.setLayout(layout2)
    tab3.setLayout(layout3)

    tabs.addTab(tab1, "Non SAP")
    tabs.addTab(tab2, "SAP")
    tabs.addTab(tab3, "JE Line Number/JE Number")

    layout.addWidget(tabs)

    ### 배치 - 탭 0
    sublayout7.addWidget(labelJE_Line, 0, 0)
    sublayout7.addWidget(self.D5_JE_Line, 0, 1)
    sublayout7.addWidget(labelJE_Number, 1, 0)
    sublayout7.addWidget(self.D5_JE_Number, 1, 1)

    layout3.addLayout(sublayout7, stretch=4)
    layout3.addLayout(sublayout8, stretch=4)
    layout3.addLayout(sublayout9, stretch=1)

    sublayout9.addStretch(2)
    sublayout9.addWidget(self.btnSave)
    sublayout9.addWidget(self.btnSaveProceed)

    ### 배치 - 탭 1
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
