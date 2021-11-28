#!/usr/bin/env python
# coding: utf-8

# In[ ]:


def Dialog12(self):  # 시나리오 11, 12, Cursor문 합본
        self.dialog12 = QDialog()
        self.dialog12.setStyleSheet('background-color: #2E2E38')
        self.dialog12.setWindowIcon(QIcon('./EY_logo.png'))

        # 시나리오12
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
        self.D12_Cost.setPlaceholderText('중요성 금액을 입력하세요')
        self.D12_Cost.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        sublayout1 = QGridLayout()
        sublayout1.addWidget(labelAccount, 0, 0)
        sublayout1.addWidget(self.new_tree, 0, 1)
        sublayout1.addWidget(labelCost, 1, 0)
        sublayout1.addWidget(self.D12_Cost, 1, 1)

        sublayout2 = QHBoxLayout()
        sublayout2.addStretch()
        sublayout2.addStretch()
        sublayout2.addWidget(self.btn)
        sublayout2.addWidget(self.btnDialog)

        main_layout1 = QVBoxLayout()
        main_layout1.addLayout(sublayout1)
        main_layout1.addStretch()
        main_layout1.addLayout(sublayout2)
        
        #시나리오11
        cursor1 = self.cnxn.cursor()

        sql1 = '''
                         SELECT 											
                                *
                         FROM  [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											

                    '''.format(field=self.selected_project_id)

        accountsname1 = pd.read_sql(sql, self.cnxn)

        self.new_tree1 = Form(self)
        self.new_tree2 = Form(self)

        self.new_tree1.tree.clear()
        self.new_tree2.tree.clear()

        for n, i in enumerate(accountsname1.AccountType.unique()):
            self.new_tree1.parent = QTreeWidgetItem(self.new_tree1.tree)

            self.new_tree1.parent.setText(0, "{}".format(i))
            self.new_tree1.parent.setFlags(self.new_tree1.parent.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
            child_items1 = accountsname1.AccountSubType[
                accountsname1.AccountType == accountsname1.AccountType.unique()[n]].unique()
            for m, x in enumerate(child_items1):
                self.new_tree1.child = QTreeWidgetItem(self.new_tree1.parent)

                self.new_tree1.child.setText(0, "{}".format(x))
                self.new_tree1.child.setFlags(self.new_tree1.child.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                grandchild_items1 = accountsname1.AccountClass[accountsname1.AccountSubType == child_items1[m]].unique()
                for o, y in enumerate(grandchild_items1):
                    self.new_tree1.grandchild = QTreeWidgetItem(self.new_tree1.child)
                    self.new_tree1.grandchild.setText(0, "{}".format(y))
                    self.new_tree1.grandchild.setFlags(
                        self.new_tree1.grandchild.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                    num_name1 = accountsname1[accountsname.AccountClass == grandchild_items1[o]].iloc[:, 2:4]
                    full_name1 = num_name1["GLAccountNumber"].map(str) + ' ' + num_name1["GLAccountName"]
                    for z in full_name1:
                        self.new_tree1.grandgrandchild = QTreeWidgetItem(self.new_tree1.grandchild)

                        self.new_tree1.grandgrandchild.setText(0, "{}".format(z))
                        self.new_tree1.grandgrandchild.setFlags(self.new_tree1.grandgrandchild.flags() | Qt.ItemIsUserCheckable)
                        self.new_tree1.grandgrandchild.setCheckState(0, Qt.Checked)
        
        self.new_tree1.get_selected_leaves()
        
        cursor2 = self.cnxn.cursor()

        sql2 = '''
                         SELECT 											
                                *
                         FROM  [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] COA											

                    '''.format(field=self.selected_project_id)

        accountsname2 = pd.read_sql(sql2, self.cnxn)
        
        for n, i in enumerate(accountsname2.AccountType.unique()):
            self.new_tree2.parent = QTreeWidgetItem(self.new_tree2.tree)
            self.new_tree2.parent.setText(0, "{}".format(i))
            self.new_tree2.parent.setFlags(self.new_tree2.parent.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
            child_items2 = accountsname2.AccountSubType[
                accountsname2.AccountType == accountsname2.AccountType.unique()[n]].unique()
            for m, x in enumerate(child_items2):
                self.new_tree2.child = QTreeWidgetItem(self.new_tree2.parent)
                self.new_tree2.child.setText(0, "{}".format(x))
                self.new_tree2.child.setFlags(self.new_tree2.child.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                grandchild_items2 = accountsname2.AccountClass[accountsname2.AccountSubType == child_items2[m]].unique()
                for o, y in enumerate(grandchild_items2):
                    self.new_tree2.grandchild = QTreeWidgetItem(self.new_tree2.child)
                    self.new_tree2.grandchild.setText(0, "{}".format(y))
                    self.new_tree2.grandchild.setFlags(
                        self.new_tree2.grandchild.flags() | Qt.ItemIsTristate | Qt.ItemIsUserCheckable)
                    num_name2 = accountsname2[accountsname.AccountClass == grandchild_items2[o]].iloc[:, 2:4]
                    full_name2 = num_name2["GLAccountNumber"].map(str) + ' ' + num_name2["GLAccountName"]
                    for z in full_name2:
                        self.new_tree2.grandgrandchild = QTreeWidgetItem(self.new_tree2.grandchild)
                        self.new_tree2.grandgrandchild.setText(0, "{}".format(z))
                        self.new_tree2.grandgrandchild.setFlags(self.new_tree2.grandgrandchild.flags() | Qt.ItemIsUserCheckable)
                        self.new_tree2.grandgrandchild.setCheckState(0, Qt.Checked)
        
        self.new_tree2.get_selected_leaves()
        
        self.btn1 = QPushButton('   Extract Data', self.dialog12)
        self.btn1.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btn1.clicked.connect(self.extButtonClicked12)
        font9 = self.btn1.font()
        font9.setBold(True)
        self.btn1.setFont(font9)

        self.btnDialog1 = QPushButton("   Close", self.dialog12)
        self.btnDialog1.setStyleSheet('color:white;  background-image : url(./bar.png)')
        self.btnDialog1.clicked.connect(self.dialog_close12)
        font10 = self.btnDialog1.font()
        font10.setBold(True)
        self.btnDialog1.setFont(font10)
        self.btn1.resize(110, 30)
        self.btnDialog1.resize(110, 30)

        # JE Line Number / JE Number 라디오 버튼
        self.rbtn3 = QRadioButton('JE Line', self.dialog12)
        self.rbtn3.setStyleSheet("color: white;")
        font11 = self.rbtn3.font()
        font11.setBold(True)
        self.rbtn3.setFont(font11)
        self.rbtn3.setChecked(True)
        self.rbtn4 = QRadioButton('JE', self.dialog12)
        self.rbtn4.setStyleSheet("color: white;")
        font12 = self.rbtn4.font()
        font12.setBold(True)
        self.rbtn4.setFont(font12)

        labelAccount1 = QLabel('A 계정명/계정 코드* : ', self.dialog12)
        labelAccount1.setStyleSheet("color: white;")
        font3 = labelAccount1.font()
        font3.setBold(True)
        labelAccount1.setFont(font3)
        
        labelAccount2 = QLabel('B 계정명/계정 코드* : ', self.dialog12)
        labelAccount2.setStyleSheet("color: white;")
        font3 = labelAccount2.font()
        font3.setBold(True)
        labelAccount2.setFont(font3)

        labelCost1 = QLabel('중요성 금액 : ', self.dialog12)
        labelCost1.setStyleSheet("color: white;")
        font3 = labelCost1.font()
        font3.setBold(True)
        labelCost1.setFont(font3)

        self.D12_Cost1 = QLineEdit(self.dialog12)
        self.D12_Cost1.setStyleSheet("background-color: white;")
        self.D12_Cost1.setPlaceholderText('중요성 금액을 입력하세요')
        self.D12_Cost1.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        sublayout3 = QGridLayout()
        sublayout3.addWidget(self.rbtn3, 0, 0)
        sublayout3.addWidget(self.rbtn4, 0, 1)
        sublayout3.addWidget(labelAccount1, 1, 0)
        sublayout3.addWidget(self.new_tree1, 1, 1)
        sublayout3.addWidget(labelAccount2, 2, 0)
        sublayout3.addWidget(self.new_tree2, 2, 1)
        sublayout3.addWidget(labelCost1, 3, 0)
        sublayout3.addWidget(self.D12_Cost1, 3, 1)

        sublayout4 = QHBoxLayout()
        sublayout4.addStretch()
        sublayout4.addStretch()
        sublayout4.addWidget(self.btn1)
        sublayout4.addWidget(self.btnDialog1)

        main_layout3 = QVBoxLayout()
        main_layout3.addLayout(sublayout3)
        main_layout3.addStretch()
        main_layout3.addLayout(sublayout4)
        
        #Cursor문
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
        
        labelCursor = QLabel('Cursor 조건 : ', self.dialog12)
        labelCursor.setStyleSheet("color: white;")
        font3 = labelCursor.font()
        font3.setBold(True)
        labelCursor.setFont(font3)

        self.cursorCondition = QTextEdit(self.dialog12)
        self.cursorCondition.setStyleSheet("background-color: white;")
        self.cursorCondition.setPlaceholderText('커서문 조건을 입력하세요')
        self.cursorCondition.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # LineEdit만 창 크기에 따라 확대/축소

        sublayout5 = QGridLayout()
        sublayout5.addWidget(labelCursor, 0, 0)
        sublayout5.addWidget(self.cursorCondition, 0, 1)

        sublayout6 = QHBoxLayout()
        sublayout6.addStretch()
        sublayout6.addStretch()
        sublayout6.addWidget(self.btn2)
        sublayout6.addWidget(self.btnDialog2)

        main_layout2 = QVBoxLayout()
        main_layout2.addLayout(sublayout5)
        main_layout2.addStretch()
        main_layout2.addLayout(sublayout6)
        
        #탭 지정
        layout = QVBoxLayout()
        tabs = QTabWidget()
        tab1 = QWidget() #시나리오12
        tab2 = QWidget() #cursor문
        tab3 = QWidget() #시나리오11
        tab1.setLayout(main_layout1)
        tab2.setLayout(main_layout2)
        tab3.setLayout(main_layout3)
        tabs.addTab(tab1, "시나리오12")
        tabs.addTab(tab2, "Cursor문")
        tabs.addTab(tab3, "시나리오11")
        layout.addWidget(tabs)

        self.dialog12.setLayout(layout)
        self.dialog12.resize(500, 300)

        # ? 제거
        self.dialog12.setWindowFlags(Qt.WindowCloseButtonHint)

        self.dialog12.setWindowTitle('Scenario')
        self.dialog12.setWindowModality(Qt.NonModal)
        self.dialog12.show()

