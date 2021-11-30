def extButtonClicked5_Non_SAP(self):

    tempSheet_NonSAP = self.D5_Sheet2.text() # 필수값
    tempYear_NonSAP = int(pname_year) # 필수값
    temp_Code_Non_SAP = self.MyInput.toPlainText() # 필수값 (계정코드)

    # temp_Code_Non_SAP = re.sub(r"[:,|\s]", ",", temp_Code_Non_SAP)
    # temp_Code_Non_SAP = re.split(",", temp_Code_Non_SAP)
    # print(temp_Code_Non_SAP) # ['447102', '445101', '289301', '289310', '289311', '289312', '289313', '289314']
    temp_Code_Non_SAP = str(temp_Code_Non_SAP)

    ### 예외처리 1 - 필수값 입력 누락
    if temp_Code_Non_SAP == '' or tempSheet_NonSAP == '' or checked_account == 'AND JournalEntries.GLAccountNumber IN ()':
        self.alertbox_open()

    ### 예외처리 2 - 시트명 중복 확인
    elif self.rbtn1.isChecked() and self.combo_sheet.findText(tempSheet_NonSAP + '_Result') != -1:
        self.alertbox_open5()

    elif self.rbtn2.isChecked() and self.combo_sheet.findText(tempSheet_NonSAP + '_Journals') != -1:
        self.alertbox_open5()

    ### 쿼리 연동
    else:

        checked_account5_NonSAP = checked_account
        cursor = self.cnxn.cursor()
        ### JE Line
        if self.rbtn1.isChecked():

            sql_query = """
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
                                    AND JournalEntries.GLAccountNumber IN ({CODE})
                                    {Account}
                            ORDER BY JournalEntries.JENumber, JournalEntries.JELineNumber  

                        """.format(field=self.selected_project_id, CODE=temp_Code_Non_SAP, Account=checked_account5_NonSAP)
        ### JE
        elif self.rbtn2.isChecked():

            sql_query = '''
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
                                [{field}_Import_CY_01].[dbo].[pbcChartOfAccounts] AS COA
                            WHERE JournalEntries.GLAccountNumber = CoA.GLAccountNumber AND JournalEntries.JENumber IN 
                            (
                                SELECT DISTINCT JENumber
                                FROM [{field}_Import_CY_01].[dbo].[pbcJournalEntries]
                                WHERE JournalEntries.GLAccountNumber IN ({CODE})
                                {Account}
                            )
                            ORDER BY JournalEntries.JENumber, JournalEntries.JELineNumber
                        '''.format(field=self.selected_project_id, CODE=temp_Code_Non_SAP, Account=checked_account5_NonSAP)

        self.dataframe = pd.read_sql(sql_query, self.cnxn)

    ### 예외처리 3 - 최대 출력 라인 초과
    if len(self.dataframe) > 1048576:
        self.alertbox_open3()

    ### 예외처리 4 - 데이터 미추출
    elif len(self.dataframe) == 0:
        self.dataframe = pd.DataFrame({'No Data': ['[연도: ' + str(tempYear_NonSAP) + ','
                                                   + '계정코드: ' + str(temp_Code_Non_SAP) + ','
                                                   + '] 라인수 ' + str(len(self.dataframe)) + '개 입니다']})

        model = DataFrameModel(self.dataframe)
        self.viewtable.setModel(model)

        ### JE Line
        if self.rbtn1.isChecked():
            self.scenario_dic[tempSheet_NonSAP + '_Result'] = self.dataframe

            key_list = list(self.scenario_dic.keys())
            result = [key_list[0], key_list[-1]]
            self.combo_sheet.addItem(str(result[1]))

            buttonReply = QMessageBox.information(self, '라인수 추출', '-당기('
                                                  + str(tempYear_NonSAP) + ')에 생성된 계정 리스트가 '
                                                  + str(len(self.dataframe) - 1)
                                                  + '건 추출되었습니다. <br> - 계정코드(' + str(temp_Code_Non_SAP)
                                                  + ')를 적용하였습니다. 추가 필터링이 필요해보입니다. <br> [전표라인번호 기준]'
                                                  , QMessageBox.Yes)
        ### JE
        elif self.rbtn2.isChecked():
            self.scenario_dic[tempSheet_NonSAP + 'Journals'] = self.dataframe

            key_list = list(self.scenario_dic.keys())
            result = [key_list[0], key_list[-1]]
            self.combo_sheet.addItem(str(result[1]))

            buttonReply = QMessageBox.information(self, '라인수 추출', '-당기('
                                                  + str(tempYear_NonSAP) + ')에 생성된 계정 리스트가 '
                                                  + str(len(self.dataframe) - 1)
                                                  + '건 추출되었습니다. <br> - 계정코드(' + str(temp_Code_Non_SAP)
                                                  + ')를 적용하였습니다. 추가 필터링이 필요해보입니다. <br> [전표라인번호 기준]'
                                                  , QMessageBox.Yes)

        if buttonReply == QMessageBox.Yes:
            self.dialog5.activateWindow()

    else:
        if self.rbtn1.isChecked():
            self.scenario_dic[tempSheet_NonSAP + '_Result'] = self.dataframe

            key_list = list(self.scenario_dic.keys())
            result = [key_list[0], key_list[-1]]
            self.combo_sheet.addItem(str(result[1]))

            model = DataFrameModel(self.dataframe)
            self.viewtable.setModel(model)

            if len(self.dataframe) - 1 <= 500:
                buttonReply = QMessageBox.information(self, '라인수 추출',
                                                      '- 당기(' + str(tempYear_NonSAP) + ')에 생성된 계정 리스트가 '
                                                      + str(len(self.dataframe) - 1) + '건 추출되었습니다. <br> - 계정코드('
                                                      + str(temp_Code_Non_SAP) + ')를 적용하였습니다. 추가 필터링이 필요해보입니다. <br> [전표라인번호 기준]'
                                                      , QMessageBox.Yes)

            else:
                buttonReply = QMessageBox.information(self, '라인수 추출',
                                                      '- 당기(' + str(tempYear_NonSAP) + ')에 생성된 계정 리스트가 '
                                                      + str(len(self.dataframe) - 1) + '건 추출되었습니다. <br> - 계정코드('
                                                      + str(temp_Code_Non_SAP) + ')를 적용하였습니다. <br> [전표라인번호 기준]'
                                                      , QMessageBox.Yes)

        elif self.rbtn2.isChecked():
            ### 시트 콤보박스에 저장
            self.scenario_dic[tempSheet_NonSAP + '_Journals'] = self.dataframe
            key_list = list(self.scenario_dic.keys())
            result = [key_list[0], key_list[-1]]
            self.combo_sheet.addItem(str(result[1]))

            model = DataFrameModel(self.dataframe)
            self.viewtable.setModel(model)

            if len(self.dataframe) - 1 <= 500:
                buttonReply = QMessageBox.information(self, '라인수 추출',
                                                      '- 당기(' + str(tempYear_NonSAP) + ')에 생성된 계정 리스트가 '
                                                      + str(len(self.dataframe) - 1) + '건 추출되었습니다. <br> - 계정코드('
                                                      + str(
                                                          temp_Code_Non_SAP) + ')를 적용하였습니다. 추가 필터링이 필요해보입니다. <br> [전표라인번호 기준]'
                                                      , QMessageBox.Yes)

            else:
                buttonReply = QMessageBox.information(self, '라인수 추출',
                                                      '- 당기(' + str(tempYear_NonSAP) + ')에 생성된 계정 리스트가 '
                                                      + str(len(self.dataframe) - 1) + '건 추출되었습니다. <br> - 계정코드('
                                                      + str(temp_Code_Non_SAP) + ')를 적용하였습니다. <br> [전표라인번호 기준]'
                                                      , QMessageBox.Yes)
        if buttonReply == QMessageBox.Yes:
            self.dialog5.activateWindow()
