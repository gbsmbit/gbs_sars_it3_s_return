Imports System.IO
Imports System.Data.SqlClient

Module SARSIT3Module

    Public Exporting As Boolean
    Public Abort As Boolean
    Public StartTime As Date

    Private Const strConnection = "Data Source=DBServerMain; " & _
                                  "Initial Catalog=GBS; " & _
                                  "User ID=adminGBS; " & _
                                  "Password=a"

    Private Const FileTypeAndMode = "IT3EXTRS.R"
    Private Const TakeOnYearEnd = "20110228"

    Dim SARSForm As New SARSIT3Form()

    Private SARSFile As String
    Private Const MaxFileRec = 10000
    Private FileSeq As Integer
    Private RecSeq As Integer
    Private TotalRecords As Integer
    Private Skipped As Integer

    Sub Main()

        SARSForm.ShowDialog()
    End Sub

    Private Function FixNull(ByVal strValue As Object) As String

        If strValue Is Nothing OrElse IsDBNull(strValue) Then
            FixNull = ""
        Else
            FixNull = strValue
        End If
    End Function

    Private Function FixNullNum(ByVal numValue As Object) As Double

        If numValue Is Nothing OrElse IsDBNull(numValue) Then
            FixNullNum = 0
        Else
            FixNullNum = numValue
        End If
    End Function

    Public Function FixedLengthString(ByVal InputString As Object, ByVal StrLen As Single, ByVal strJustify As String, ByVal strPad As Char) As String

        Dim OutputString As String

        InputString = FixNull(InputString)
        If InputString.Length > StrLen Then
            InputString = InputString.Substring(0, StrLen)
        End If
        While InputString.Length < StrLen
            If strJustify = "Left" Then
                InputString = InputString & strPad
            Else
                InputString = strPad & InputString
            End If
        End While
        FixedLengthString = InputString
    End Function

    Private Function GetAddress(ByVal CIFId As Integer, ByVal AddressType As String) As Array

        Dim Address(5) As String

        Dim strSQL As String
        If AddressType = "Postal" Then
            strSQL = "SELECT TOP 1 * FROM tblAddress INNER JOIN mblCity ON tblAddress.cityId = mblCity.cityId " & _
                     "WHERE tblAddress.cifId = " & CIFId & " AND tblAddress.adtid = 2"
        Else
            strSQL = "SELECT TOP 1 * FROM tblAddress INNER JOIN mblCity ON tblAddress.cityId = mblCity.cityId " & _
                     "WHERE tblAddress.cifId = " & CIFId & " AND " & _
                     "(tblAddress.adtid = 1 OR tblAddress.adtid = 14 OR tblAddress.adtid = 16)"
        End If
        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd = New SqlCommand(strSQL, cn)

        cn.Open()
        drSQL = cmd.ExecuteReader()
        While drSQL.Read()
            Address(1) = FixNull(drSQL.Item("addrLine1"))
            Address(2) = FixNull(drSQL.Item("addrLine2"))
            Address(3) = FixNull(drSQL.Item("addrLine3"))
            Address(4) = FixNull(drSQL.Item("cityDesc"))
            Address(5) = FixNull(drSQL.Item("addrCode"))
        End While
        cn.Close()
        Address(1) = FixedLengthString(Address(1), 35, "Left", " ")
        Address(2) = FixedLengthString(Address(2), 35, "Left", " ")
        Address(3) = FixedLengthString(Address(3), 35, "Left", " ")
        Address(4) = FixedLengthString(Address(4), 35, "Left", " ")
        Address(5) = FixedLengthString(Address(5), 10, "Left", " ")
        GetAddress = Address
    End Function

    Public Function FileHeader()

        Const SEC_ID = "H"
        Const INFO_TYPE = "IT3EXTRS"
        Const INFO_SUBTYPE = " "
        Const FILE_SERIES_CTL = "S"
        Const EXT_SYS = "GBSMUTUA"
        Const VER_NO = "1"

        Dim GEN_TIME As String
        Dim Now As DateTime = DateTime.Now

        GEN_TIME = Now.Year & FixedLengthString(Now.Month, 2, "Right", "0") & FixedLengthString(Now.Day, 2, "Right", "0") & _
                   FixedLengthString(Now.Hour, 2, "Right", "0") & FixedLengthString(Now.Minute, 2, "Right", "0") & FixedLengthString(Now.Second, 2, "Right", "0")

        FileHeader = SEC_ID & INFO_TYPE & FixedLengthString(INFO_SUBTYPE, 8, "Left", " ") & _
                     SARSForm.TestFile & FILE_SERIES_CTL & EXT_SYS & FixedLengthString(VER_NO, 8, "Right", " ") & _
                     FixedLengthString(SARSForm.Reference, 14, "Right", " ") & GEN_TIME
    End Function

    Private Function PersonalDetailsRecord(ByVal CIFInfo As Object) As String

        Const SEC_ID = "P"

        Dim IT3_PERS_ID As String
        Dim IT_REF_NO As String
        Dim PERIOD_START As String = SARSForm.PeriodStart
        Dim PERIOD_END As String = SARSForm.PeriodEnd
        Dim TP_CATEGORY As String
        Dim TP_ID As String
        Dim TP_OTHER_ID As String
        Dim CO_REG_NO As String
        Dim TRUST_DEED_NO As String
        Dim TP_NAME As String
        Dim TP_INITS As String
        Dim TP_FIRSTNAMES As String
        Dim TP_DOB As String
        Dim TP_TRADE_NAME As String
        Dim TP_POST_ADDR(5) As String
        'Dim TP_POST_CODE As String
        Dim TP_PHY_ADDR(5) As String
        'Dim TP_PHY_CODE As String
        Dim TP_SA_RES As String
        Dim PARTNERSHIP As String

        IT3_PERS_ID = FixedLengthString(CIFInfo(4), 25, "Left", " ")
        TP_POST_ADDR = GetAddress(CIFInfo(0), "Postal")
        TP_PHY_ADDR = GetAddress(CIFInfo(0), "Physical")
        PARTNERSHIP = "N"

        Select Case CIFInfo(1)

            Case 1   'Individuals
                IT_REF_NO = FixedLengthString(CIFInfo(21), 10, "Left", " ")
                TP_CATEGORY = "01"
                TP_ID = FixedLengthString(CIFInfo(16), 13, "Left", " ")
                If FixNull(CIFInfo(23)) <> "" Then
                    TP_OTHER_ID = FixedLengthString(CIFInfo(23), 10, "Left", " ")
                Else
                    If FixNull(CIFInfo(24)) <> "" Then
                        TP_OTHER_ID = FixedLengthString(CIFInfo(24), 10, "Left", " ")
                    Else
                        TP_OTHER_ID = FixedLengthString(" ", 10, "Left", " ")
                    End If
                End If
                CO_REG_NO = FixedLengthString(" ", 15, "Left", " ")
                TRUST_DEED_NO = FixedLengthString(" ", 10, "Left", " ")
                TP_NAME = FixedLengthString(CIFInfo(18), 120, "Left", " ")
                TP_INITS = FixedLengthString(CIFInfo(19), 5, "Left", " ")
                TP_FIRSTNAMES = FixedLengthString(CIFInfo(17), 90, "Left", " ")
                TP_DOB = FixedLengthString(CIFInfo(22), 8, "Left", " ")
                TP_TRADE_NAME = FixedLengthString(" ", 120, "Left", " ")
                If FixNull(CIFInfo(12)) = "" Then
                    TP_SA_RES = "Y"
                Else
                    If CDbl(FixNull(CIFInfo(12))) = 510 Then
                        TP_SA_RES = "Y"
                    Else
                        TP_SA_RES = "N"
                    End If
                End If
            Case 2   'Companies
                IT_REF_NO = FixedLengthString(CIFInfo(45), 10, "Left", " ")
                Select Case CIFInfo(40)
                    Case 12, 13, 27
                        TP_CATEGORY = "02"
                    Case 18, 19, 26
                        TP_CATEGORY = "03"
                    Case Else
                        TP_CATEGORY = "00"
                End Select
                TP_ID = FixedLengthString(" ", 13, "Left", " ")
                TP_OTHER_ID = FixedLengthString(" ", 10, "Left", " ")
                If TP_CATEGORY = "03" Then
                    CO_REG_NO = FixedLengthString(" ", 15, "Left", " ")
                    TRUST_DEED_NO = FixedLengthString(CIFInfo(41), 10, "Left", " ")
                Else
                    CO_REG_NO = FixedLengthString(CIFInfo(41), 15, "Left", " ")
                    TRUST_DEED_NO = FixedLengthString(" ", 10, "Left", " ")
                End If
                TP_NAME = FixedLengthString(CIFInfo(43), 120, "Left", " ")
                TP_INITS = FixedLengthString(" ", 5, "Left", " ")
                TP_FIRSTNAMES = FixedLengthString(" ", 90, "Left", " ")
                TP_DOB = FixedLengthString(" ", 8, "Left", " ")
                TP_TRADE_NAME = FixedLengthString(CIFInfo(44), 120, "Left", " ")
                If CIFInfo(46) = 510 Then
                    TP_SA_RES = "Y"
                Else
                    TP_SA_RES = "N"
                End If
        End Select

        PersonalDetailsRecord = SEC_ID & _
                                IT3_PERS_ID & _
                                IT_REF_NO & _
                                PERIOD_START & _
                                PERIOD_END & _
                                TP_CATEGORY & _
                                TP_ID & _
                                TP_OTHER_ID & _
                                CO_REG_NO & _
                                TRUST_DEED_NO & _
                                TP_NAME & _
                                TP_INITS & _
                                TP_FIRSTNAMES & _
                                TP_DOB & _
                                TP_TRADE_NAME & _
                                TP_POST_ADDR(1) & _
                                TP_POST_ADDR(2) & _
                                TP_POST_ADDR(3) & _
                                TP_POST_ADDR(4) & _
                                TP_POST_ADDR(5) & _
                                TP_PHY_ADDR(1) & _
                                TP_PHY_ADDR(2) & _
                                TP_PHY_ADDR(3) & _
                                TP_PHY_ADDR(4) & _
                                TP_PHY_ADDR(5) & _
                                TP_SA_RES & _
                                PARTNERSHIP
    End Function

    Private Function AccountOpeningBalance(ByVal accId As Integer) As Integer

        Dim strSQL As String = "SELECT TOP (1) trnAccountBal " & _
                               "FROM tblTransactions " & _
                               "WHERE (accId = " & accId & ") " & _
                               "ORDER BY trnId ASC"

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd = New SqlCommand(strSQL, cn)

        cn.Open()
        drSQL = cmd.ExecuteReader()
        drSQL.Read()
        AccountOpeningBalance = FixNullNum(drSQL.Item(0)) * 100
        cn.Close()
    End Function

    Private Function AccountClosingBalance(ByVal accId As Integer) As Integer

        Dim strSQL As String = "SELECT TOP (1) trnAccountBal " & _
                               "FROM tblTransactions " & _
                               "WHERE (accId = " & accid & ") " & _
                               "ORDER BY trnId ASC"

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd = New SqlCommand(strSQL, cn)

        cn.Open()
        drSQL = cmd.ExecuteReader()
        drSQL.Read()
        AccountClosingBalance = FixNullNum(drSQL.Item(0)) * 100
        cn.Close()
    End Function

    Private Function AccountCloseDate(ByVal accId As Integer) As String

        Dim strSQL As String = "SELECT TOP (1) trnTransactionDate " & _
                               "FROM tblTransactions " & _
                               "WHERE (accId = " & accId & ") " & _
                               "ORDER BY trnId DESC"

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd = New SqlCommand(strSQL, cn)

        cn.Open()
        drSQL = cmd.ExecuteReader()
        drSQL.Read()
        AccountCloseDate = FixNull(drSQL.Item(0))
        cn.Close()
    End Function

    Private Function IncomeDetails(ByVal AccountInfo As Object) As String

        Const SEC_ID = "I"
        Const INCOME_NATURE = "4201"
        Const BRANCH_CODE = "000000"

        Dim IT3_PERS_ID As String
        Dim INCOME_PAID As String
        Dim ACCOUNT_NO As String
        Dim ACCOUNT_TYPE As String
        Dim START_DATE As String
        Dim START_BAL As String
        Dim START_BAL_SIGN As String
        Dim END_DATE As String
        Dim END_BAL As String
        Dim END_BAL_SIGN As String
        Dim FOREIGN_TAX_PAID As String

        Dim OpeningBalance As Integer
        Dim ClosingBalance As Integer
        Dim CloseDate As String

        IT3_PERS_ID = FixedLengthString(AccountInfo(0), 25, "Left", " ")
        INCOME_PAID = FixedLengthString(FixNullNum((AccountInfo(1)) + FixNullNum(AccountInfo(2))) * 100, 15, "Right", " ")
        ACCOUNT_NO = FixedLengthString(AccountInfo(3), 20, "Left", " ")
        Select Case AccountInfo(4)
            Case 11, 12, 14, 16
                ACCOUNT_TYPE = "03"
            Case 13, 18, 19
                ACCOUNT_TYPE = "06"
            Case 51
                ACCOUNT_TYPE = "05"
            Case Else
                ACCOUNT_TYPE = "17"
        End Select
        If AccountInfo(5) > SARSForm.PeriodStart Then
            OpeningBalance = AccountOpeningBalance(AccountInfo(7))
            START_DATE = AccountInfo(5)
            START_BAL = FixedLengthString(OpeningBalance, 15, "Right", " ")
            If OpeningBalance >= 0 Then
                START_BAL_SIGN = "C"
            Else
                START_BAL_SIGN = "D"
            End If
        Else
            START_DATE = SARSForm.PeriodStart
            START_BAL = FixedLengthString(" ", 15, "Right", " ")
            START_BAL_SIGN = "C"
        End If
        If AccountInfo(6) > 0 Then
            CloseDate = AccountCloseDate(AccountInfo(7))
        Else
            CloseDate = "0"
        End If
        If CloseDate > 0 And CloseDate < SARSForm.PeriodEnd Then
            END_DATE = CloseDate
            END_BAL = FixedLengthString(" ", 15, "Right", " ")
            END_BAL_SIGN = "C"
        Else
            ClosingBalance = AccountClosingBalance(AccountInfo(7))
            END_DATE = SARSForm.PeriodEnd
            END_BAL = FixedLengthString(ClosingBalance, 15, "Right", " ")
            If ClosingBalance >= 0 Then
                END_BAL_SIGN = "C"
            Else
                END_BAL_SIGN = "D"
            End If
        End If
        FOREIGN_TAX_PAID = FixedLengthString(" ", 15, "Right", " ")

        IncomeDetails = SEC_ID & _
                        IT3_PERS_ID & _
                        INCOME_NATURE & _
                        INCOME_PAID & _
                        ACCOUNT_NO & _
                        BRANCH_CODE & _
                        ACCOUNT_TYPE & _
                        START_DATE & _
                        START_BAL & _
                        START_BAL_SIGN & _
                        END_DATE & _
                        END_BAL & _
                        END_BAL_SIGN & _
                        FOREIGN_TAX_PAID

    End Function

    Private Function FileTrailer(ByVal NoOfRecords As Integer) As String

        FileTrailer = "T" & FixedLengthString(NoOfRecords, 8, "Right", "0")
    End Function

    Private Sub WriteIncomeRecords(ByVal cifId As Integer)

        Dim strSQL As String = "SELECT NotImported.cifNO, NotImported.niAmnt, Imported.iAmnt, NotImported.accNO, NotImported.AccType, NotImported.OpenDate, NotImported.Status, NotImported.accId " & _
                               "FROM (SELECT tblCIF.cifNO as cifNO, sum(tblAccountMemoInterest.accIntAmt) as niAmnt, tblAccount.accNO as accNO, tblAccount.atId as AccType, tblAccount.accRegDate as OpenDate, tblAccount.asId as Status, tblAccount.accId as accId " & _
                                     "FROM tblCIF INNER JOIN jblCIFAccount ON tblCIF.cifId = jblCIFAccount.cifId " & _
                                                 "INNER JOIN tblAccount ON jblCIFAccount.accId = tblAccount.accId " & _
                                                 "INNER JOIN tblAccountMemoInterest ON tblAccount.accId = tblAccountMemoInterest.accId " & _
                                     "WHERE (tblCIF.cifId = " & cifId & ") AND (tblAccount.atId < 71) AND " & _
                                           "(tblAccountMemoInterest.accIntCapd = 1) AND " & _
                                           "(tblAccountMemoInterest.accPostDate >= '20100101') AND " & _
                                           "(tblAccountMemoInterest.accPostDate <= '20101231') AND " & _
                                           "(NOT (tblAccountMemoInterest.accIntRef LIKE '%IMPORT FROM SA THRIFT%')) " & _
                                     "GROUP BY tblCIF.cifNO,tblAccount.accNO, tblAccount.atId, tblAccount.accRegDate, tblAccount.asId, tblAccount.accId) as NotImported " & _
                               "LEFT OUTER JOIN " & _
                                     "(SELECT tblCIF.cifNO as cifNO, sum(tblAccountMemoInterest.accIntAmt) as iAmnt, tblAccount.accNO as accNO " & _
                                     "FROM tblCIF INNER JOIN jblCIFAccount ON tblCIF.cifId = jblCIFAccount.cifId " & _
                                                 "INNER JOIN tblAccount ON jblCIFAccount.accId = tblAccount.accId " & _
                                                 "INNER JOIN tblAccountMemoInterest ON tblAccount.accId = tblAccountMemoInterest.accId " & _
                                     "WHERE (tblCIF.cifId = " & cifId & ") AND (tblAccount.atId < 71) AND " & _
                                           "(tblAccountMemoInterest.accIntCapd = 1) AND " & _
                                           "(tblAccountMemoInterest.accIntRef = 'IMPORT FROM SA THRIFT') " & _
                                     "GROUP BY tblCIF.cifNO,tblAccount.accNO) AS Imported " & _
                               "ON NotImported.accNO = Imported.accNO"

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd = New SqlCommand(strSQL, cn)

        cn.Open()
        drSQL = cmd.ExecuteReader()
        While drSQL.Read()

            Dim AccountInfo(drSQL.FieldCount - 1) As Object
            Dim fieldCount As Integer = drSQL.GetValues(AccountInfo)
            WriteToFile(SARSFile, IncomeDetails(AccountInfo))
            TotalRecords = TotalRecords + 1
        End While
        cn.Close()
    End Sub

    Private Function GetNoOfClientAccounts(ByVal cifId As Integer) As Integer

        Dim strSQL As String = "SELECT * FROM " & _
                                "(SELECT tblAccount.accNO " & _
                                 "FROM jblCIFAccount INNER JOIN tblAccount ON jblCIFAccount.accId = tblAccount.accId " & _
                                                    "INNER JOIN tblAccountMemoInterest ON tblAccount.accId = tblAccountMemoInterest.accId " & _
                                 "WHERE(jblCIFAccount.cifId = " & cifId & ") And " & _
                                      "(tblAccount.atId < 71) And " & _
                                      "(tblAccountMemoInterest.accIntCapd = 1) And " & _
                                      "(tblAccountMemoInterest.accPostDate >= '" & SARSForm.PeriodStart & "') AND " & _
                                      "(tblAccountMemoInterest.accPostDate <= '" & SARSForm.PeriodEnd & " ') AND " & _
                                      "(NOT (tblAccountMemoInterest.accIntRef LIKE '%IMPORT FROM SA THRIFT%')) AND " & _
                                      "(tblAccountMemoInterest.accIntAmt <> 0) " & _
                                 "GROUP BY tblAccount.accNO) AS NotImported " & _
                               "LEFT OUTER JOIN " & _
                                "(SELECT tblAccount.accNO " & _
                                 "FROM jblCIFAccount INNER JOIN tblAccount ON jblCIFAccount.accId = tblAccount.accId " & _
                                                    "INNER JOIN tblAccountMemoInterest ON tblAccount.accId = tblAccountMemoInterest.accId " & _
                                 "WHERE (jblCIFAccount.cifId = " & cifId & ") AND " & _
                                       "(tblAccount.atId < 71) And " & _
                                       "(tblAccountMemoInterest.accIntCapd = 1) AND " & _
                                       "(tblAccountMemoInterest.accIntRef = 'IMPORT FROM SA THRIFT') AND " & _
                                       "(tblAccountMemoInterest.accIntAmt <> 0) " & _
                                 "GROUP BY tblAccount.accNO) AS Imported " & _
                               "ON Imported.accNO = NotImported.accNO"

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd = New SqlCommand(strSQL, cn)

        GetNoOfClientAccounts = 0
        cn.Open()
        drSQL = cmd.ExecuteReader()
        While drSQL.Read
            GetNoOfClientAccounts = GetNoOfClientAccounts + 1
        End While
        cn.Close()
    End Function

    Private Function GetInterestIncome(ByVal cifId As Integer) As Double

        Dim strSQL As String = "SELECT Imported.iAmnt + NotImported.niAmnt " & _
                               "FROM " & _
                                "(SELECT jblCIFAccount.cifId, SUM(tblAccountMemoInterest.accIntAmt) AS iAmnt " & _
                                 "FROM jblCIFAccount INNER JOIN tblAccount ON jblCIFAccount.accId = tblAccount.accId " & _
                                                    "INNER JOIN tblAccountMemoInterest ON tblAccount.accId = tblAccountMemoInterest.accId " & _
                                 "WHERE (jblCIFAccount.cifId = " & cifId & ") AND " & _
                                       "(tblAccount.atId < 71) AND " & _
                                       "(tblAccountMemoInterest.accIntCapd = 1) AND " & _
                                       "(tblAccountMemoInterest.accIntRef = 'IMPORT FROM SA THRIFT') AND " & _
                                       "(tblAccountMemoInterest.accIntAmt <> 0) " & _
                                 "GROUP BY jblCIFAccount.cifId ) AS Imported " & _
                               "INNER JOIN " & _
                                "(SELECT jblCIFAccount.cifId, SUM(tblAccountMemoInterest.accIntAmt) AS niAmnt " & _
                                 "FROM jblCIFAccount INNER JOIN tblAccount ON jblCIFAccount.accId = tblAccount.accId " & _
                                                    "INNER JOIN tblAccountMemoInterest ON tblAccount.accId = tblAccountMemoInterest.accId " & _
                                 "WHERE (jblCIFAccount.cifId = " & cifId & ") AND " & _
                                       "(tblAccount.atId < 71) AND " & _
                                       "(tblAccountMemoInterest.accIntCapd = 1) AND " & _
                                       "(tblAccountMemoInterest.accPostDate >= '" & SARSForm.PeriodStart & "') AND " & _
                                       "(tblAccountMemoInterest.accPostDate <= '" & SARSForm.PeriodEnd & " ') AND " & _
                                       "(NOT (tblAccountMemoInterest.accIntRef LIKE '%IMPORT FROM SA THRIFT%')) AND " & _
                                       "(tblAccountMemoInterest.accIntAmt <> 0) " & _
                                 "GROUP BY jblCIFAccount.cifId ) AS NotImported " & _
                               "ON Imported.cifId = NotImported.cifId"

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd = New SqlCommand(strSQL, cn)

        GetInterestIncome = 0
        cn.Open()
        drSQL = cmd.ExecuteReader()
        While drSQL.Read()
            GetInterestIncome = GetInterestIncome + drSQL.Item(0)
        End While
        cn.Close()
    End Function

    Public Sub CreateExportFile()

        Dim strSQL As String = "SELECT * FROM tblCIF LEFT OUTER JOIN tblCIFIndividual ON tblCIF.cifId = tblCIFIndividual.cifId " & _
                               "LEFT OUTER JOIN tblCIFCompany ON tblCIF.cifId = tblCIFCompany.cifId"
        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd = New SqlCommand(strSQL, cn)
        Dim NoOfClientAccounts As Integer
        Dim InterestIncome As Double

        StartTime = Now
        SARSForm.SetButton = "Stop"
        Exporting = True
        Abort = False
        TotalRecords = 0
        Skipped = 0
        FileSeq = 1
        RecSeq = 0
        SARSFile = SARSForm.OutputPath & FileTypeAndMode & FixedLengthString(FileSeq, 6, "Right", "0")
        If File.Exists(SARSFile) Then
            File.Delete(SARSFile)
        End If
        WriteToFile(SARSFile, FileHeader)
        cn.Open()
        drSQL = cmd.ExecuteReader()
        While drSQL.Read() And Not Abort
            NoOfClientAccounts = GetNoOfClientAccounts(drSQL.Item(0))
            InterestIncome = GetInterestIncome(drSQL.Item(0))
            If NoOfClientAccounts > 0 And InterestIncome > 0 Then
                If RecSeq > MaxFileRec - NoOfClientAccounts - 3 Then        '3 = Header + Trailer + Personal Info Record
                    WriteToFile(SARSFile, FileTrailer(RecSeq - 1))
                    FileSeq = FileSeq + 1
                    RecSeq = 0
                    SARSFile = SARSForm.OutputPath & FileTypeAndMode & FixedLengthString(FileSeq, 6, "Right", "0")
                    If File.Exists(SARSFile) Then
                        File.Delete(SARSFile)
                    End If
                    WriteToFile(SARSFile, FileHeader)
                End If
                Dim CIFInfo(drSQL.FieldCount - 1) As Object
                Dim fieldCount As Integer = drSQL.GetValues(CIFInfo)
                WriteToFile(SARSFile, PersonalDetailsRecord(CIFInfo))
                TotalRecords = TotalRecords + 1
                WriteIncomeRecords(drSQL.Item(0))
                SARSForm.ClientRecords = TotalRecords
            Else
                Skipped = Skipped + 1
                SARSForm.ClientsSkipped = Skipped
            End If
            Application.DoEvents()
        End While
        WriteToFile(SARSFile, FileTrailer(RecSeq - 1))
        cn.Close()
        Exporting = False
        SARSForm.SetButton = "Exit"
    End Sub

    Private Sub WriteToFile(ByVal SARSFile As String, ByVal StringToWrite As String)

        Dim SARSWriter As New System.IO.StreamWriter(SARSFile, True)
        SARSWriter.WriteLine(StringToWrite)
        SARSWriter.Close()
        RecSeq = RecSeq + 1
        SARSForm.FileNo = FileSeq
        SARSForm.RecordNo = RecSeq
    End Sub

End Module
