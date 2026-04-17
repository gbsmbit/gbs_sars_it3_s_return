Imports System.IO
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Text

Module SARSIT3Module

    Public Exporting As Boolean
    Public Abort As Boolean
    Public StartTime As Date

    'Private Const strConnection As String = "Data Source=DBServerMain; " & _
    '                                        "Initial Catalog=GBS; " & _
    '                                        "User ID=adminGBS; " & _
    '                                        "Password=a"

    Public strConnection As String

    'Public strDatabaseConnection As String

    Private Const FileTypeAndMode As String = "IT3EXTRS.C"
    Private Const TakeOnYearEnd As String = "20110228"

    Public IniFile As String = Application.StartupPath & "\SARSIT3.ini"
    Private CIFsToSkipFile As String = Application.StartupPath & "\CIFsToSkip.ini"
    Dim CIFsToSkip As New Collection()

    Private Const lenID As Byte = 13
    Private Const lenTaxRef As Byte = 10
    Private Const SARSOtherAccountCode As String = "17"

    Dim SARSForm As New SARSIT3Form()

    Private Institution As String
    Private RegulatorRegistrationNo As String
    Private RegulatorDesignation As String
    Private InstitutionTaxReference As String
    Private BranchCode As String
    Public DataFileLocation As String = Application.StartupPath & "\"
    Public DataFileLocationTest As String = Application.StartupPath & "\"
    Private FileDataType As String
    Private FileLayoutVersion As String
    Private FileNameDelimiter As String = "_"
    Private DataDelimiter As String = "|"
    Private SARSSourceID As String
    Private SourceSystem As String
    Private SourceSystemVersion As String
    Private ContactPersonName As String
    Private ContactPersonSurname As String
    Private Telephone1 As String
    Private Telephone2 As String
    Private CellNo As String
    Private EMail As String
    Private NoCountry As String
    Private SourceCode As String
    Public SkipInvalidRefs As Boolean
    Private NatureOfPerson As String
    Private PostalAddress1 As String
    Private PostalAddress2 As String
    Private PostalAddress3 As String
    Private PostalAddress4 As String
    Private PostalCode As String

    Public IT3sHeaderTable As String
    Private IT3sClientDataTable As String
    Private IT3sExceptionTable As String
    Private IT3sUniqueNoTable As String
    Private IT3sResponseTable As String

    Public it3sPeriod As String
    Public it3sSubmissionNo As Integer

    Private SARSFile As String
    Private MaxFileRec As Integer = 10000
    Private FileSeq As Integer
    Private RecSeq As Integer
    Private TotalRecords As Integer
    Private Skipped As Integer
    Private ClientsSkippedDueToPersonalInfo As Integer
    Private ClientsSkippedDueToReferences As Integer
    Private ClientsSkippedDueInactivity As Integer
    Private ClientsSkippedDueToList As Integer

    Private Rec_No As Integer = 0
    'Private Unique_No As Integer = 0

    Private ExportFiles(1, -1) As String

    Private TotalMoney As Double

    Private InvalidIDExceptionRaised As Boolean
    Private InvalidTaxNoExceptionRaised As Boolean
    Private InvalidCoRegNoExceptionRaised As Boolean
    Private InvalidTrustRegNoExceptionRaised As Boolean

    Sub Main()

        Dim frmLogin As New frmDatabaseLogin

        frmLogin.ShowDialog()
        If strConnection <> "" Then
            'LoadIniFile()
            LoadConstants()
            LoadCIFsToSkip()
            SARSForm.ShowDialog()
        Else
            MessageBox.Show("Database login failed!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub LoadIniFile()

        Dim strLine As String

        Try
            Dim ReadIni As New StreamReader(IniFile)

            While ReadIni.Peek <> -1
                strLine = ReadIni.ReadLine()
                Select Case UCase(strLine.Substring(0, strLine.IndexOf(":")))
                    Case "INSTITUTION"
                        Institution = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "REGULATORREGISTRATIONNO"
                        RegulatorRegistrationNo = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "REGULATORDESIGNATION"
                        RegulatorDesignation = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "BRANCHCODE"
                        BranchCode = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "TAXREF"
                        InstitutionTaxReference = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "OUTPUTPATH"
                        DataFileLocation = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "OUTPUTPATHTEST"
                        DataFileLocationTest = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "FILEDATATYPE_IT3S"
                        FileDataType = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "FILELAYOUTVERSION_IT3S"
                        FileLayoutVersion = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "FILENAMEDELIMITER"
                        FileNameDelimiter = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "DATADELIMITER"
                        DataDelimiter = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "SOURCEID"
                        SARSSourceID = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "SOURCESYSTEM"
                        SourceSystem = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "SOURCESYSTEMVERSION"
                        SourceSystemVersion = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "CONTACTPERSONNAME"
                        ContactPersonName = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "CONTACTPERSONSURNAME"
                        ContactPersonSurname = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "TELEPHONE1"
                        Telephone1 = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "TELEPHONE2"
                        Telephone2 = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "CELL"
                        CellNo = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "EMAIL"
                        EMail = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "HEADERTABLE"
                        IT3sHeaderTable = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "CLIENTDATATABLE"
                        IT3sClientDataTable = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "EXCEPTIONTABLE"
                        IT3sExceptionTable = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "UNIQUENOTABLE"
                        IT3sUniqueNoTable = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "RESPONSETABLE"
                        IT3sResponseTable = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "NOCOUNTRY"
                        NoCountry = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "SOURCECODE"
                        SourceCode = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "MAXRECORDS"
                        MaxFileRec = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "SKIPINVALIDREFS"
                        SkipInvalidRefs = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "NATUREOFPERSON"
                        NatureOfPerson = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "POSTALADDRESS1"
                        PostalAddress1 = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "POSTALADDRESS2"
                        PostalAddress2 = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "POSTALADDRESS3"
                        PostalAddress3 = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "POSTALADDRESS4"
                        PostalAddress4 = strLine.Substring(strLine.IndexOf(":") + 1)
                    Case "POSTALCODE"
                        PostalCode = strLine.Substring(strLine.IndexOf(":") + 1)
                End Select
            End While
            ReadIni.Close()
        Catch ex As Exception
            MessageBox.Show("An error occurred while reading the initialisation file", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub LoadConstants()

        Dim strSQL As String = "SELECT * " &
                               "FROM mblIT3Parameters"

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)

        Try
            cn.Open()
            drSQL = cmd.ExecuteReader()
            While drSQL.Read()
                Select Case drSQL.Item(0)
                    Case "INSTITUTION"
                        Institution = FixNull(drSQL.Item(1))
                    Case "REGULATORREGISTRATIONNO"
                        RegulatorRegistrationNo = FixNull(drSQL.Item(1))
                    Case "REGULATORDESIGNATION"
                        RegulatorDesignation = FixNull(drSQL.Item(1))
                    Case "BRANCHCODE"
                        BranchCode = FixNull(drSQL.Item(1))
                    Case "TAXREF"
                        InstitutionTaxReference = FixNull(drSQL.Item(1))
                    Case "OUTPUTPATH_IT3S"
                        DataFileLocation = FixNull(drSQL.Item(1))
                    Case "OUTPUTPATHTEST_IT3S"
                        DataFileLocationTest = FixNull(drSQL.Item(1))
                    Case "FILEDATATYPE_IT3S"
                        FileDataType = FixNull(drSQL.Item(1))
                    Case "FILELAYOUTVERSION_IT3S"
                        FileLayoutVersion = FixNull(drSQL.Item(1))
                    Case "FILENAMEDELIMITER"
                        FileNameDelimiter = FixNull(drSQL.Item(1))
                    Case "DATADELIMITER"
                        DataDelimiter = FixNull(drSQL.Item(1))
                    Case "SOURCEID"
                        SARSSourceID = FixNull(drSQL.Item(1))
                    Case "SOURCESYSTEM"
                        SourceSystem = FixNull(drSQL.Item(1))
                    Case "SOURCESYSTEMVERSION"
                        SourceSystemVersion = FixNull(drSQL.Item(1))
                    Case "CONTACTPERSONNAME"
                        ContactPersonName = FixNull(drSQL.Item(1))
                    Case "CONTACTPERSONSURNAME"
                        ContactPersonSurname = FixNull(drSQL.Item(1))
                    Case "TELEPHONE1"
                        Telephone1 = FixNull(drSQL.Item(1))
                    Case "TELEPHONE2"
                        Telephone2 = FixNull(drSQL.Item(1))
                    Case "CELL"
                        CellNo = FixNull(drSQL.Item(1))
                    Case "EMAIL"
                        EMail = FixNull(drSQL.Item(1))
                    Case "HEADERTABLE_IT3S"
                        IT3sHeaderTable = FixNull(drSQL.Item(1))
                    Case "CLIENTDATATABLE_IT3S"
                        IT3sClientDataTable = FixNull(drSQL.Item(1))
                    Case "EXCEPTIONTABLE_IT3S"
                        IT3sExceptionTable = FixNull(drSQL.Item(1))
                    Case "UNIQUENOTABLE_IT3S"
                        IT3sUniqueNoTable = FixNull(drSQL.Item(1))
                    Case "RESPONSETABLE_IT3S"
                        IT3sResponseTable = FixNull(drSQL.Item(1))
                    Case "NOCOUNTRY"
                        NoCountry = FixNull(drSQL.Item(1))
                    Case "SOURCECODE"
                        SourceCode = FixNull(drSQL.Item(1))
                    Case "MAXRECORDS"
                        MaxFileRec = FixNull(drSQL.Item(1))
                    Case "SKIPINVALIDREFS"
                        SkipInvalidRefs = FixNull(drSQL.Item(1))
                    Case "NATUREOFPERSON"
                        NatureOfPerson = FixNull(drSQL.Item(1))
                    Case "POSTALADDRESS1"
                        PostalAddress1 = FixNull(drSQL.Item(1))
                    Case "POSTALADDRESS2"
                        PostalAddress2 = FixNull(drSQL.Item(1))
                    Case "POSTALADDRESS3"
                        PostalAddress3 = FixNull(drSQL.Item(1))
                    Case "POSTALADDRESS4"
                        PostalAddress4 = FixNull(drSQL.Item(1))
                    Case "POSTALCODE"
                        PostalCode = FixNull(drSQL.Item(1))
                End Select
            End While
        Catch ex As Exception
            MessageBox.Show("Failed in loading parameters from database - going to try initialisation file", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error)
            LoadIniFile()
        End Try
        cn.Close()
    End Sub

    Private Sub LoadCIFsToSkip()

        Try
            Dim ReadCIFs As New StreamReader(CIFsToSkipFile)

            While ReadCIFs.Peek <> -1
                CIFsToSkip.Add(FixedLengthString(ReadCIFs.ReadLine(), 6, "Right", "0"))
            End While
            ReadCIFs.Close()
        Catch ex As Exception
        End Try
    End Sub

    Private Function CIFToSkip(ByVal CIF As String) As Boolean

        Dim CIFIndex As Integer
        Dim i As Integer

        i = 1
        CIFIndex = 0
        While i <= CIFsToSkip.Count And CIFIndex = 0
            If CIFsToSkip.Item(i) = CIF Then
                CIFIndex = i
            End If
            i = i + 1
        End While
        If CIFIndex > 0 Then
            CIFToSkip = True
        Else
            CIFToSkip = False
        End If
    End Function

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

    Public Function FixedLengthString(ByVal InputString As Object, ByVal StrLen As Single, ByVal strJustify As String, ByVal strPad As Char, Optional ByVal chrDrop1 As Char = Chr(0), Optional ByVal chrDrop2 As Char = Chr(0)) As String

        Dim strTmp As String = ""

        InputString = FixNull(InputString)
        If chrDrop1 <> Chr(0) Then
            InputString = DropChar(InputString, chrDrop1)
        End If
        If chrDrop2 <> Chr(0) Then
            InputString = DropChar(InputString, chrDrop2)
        End If
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
        FixedLengthString = InputString.ToString.ToUpper
    End Function

    Private Function DropChar(ByVal InputString As String, ByVal chrDrop As Char) As String

        Return InputString.Replace(chrDrop, "")
    End Function

    'Private Function DropChar(ByVal InputString As String, ByVal chrDrop As Char) As String

    '    Dim i As Integer
    '    Dim strTmp As String = ""

    '    For i = 0 To InputString.Length - 1
    '        If InputString.Substring(i, 1) <> chrDrop Then
    '            strTmp = strTmp & InputString.Substring(i, 1)
    '        End If
    '    Next
    '    Return strTmp
    'End Function

    Private Function DateWithDashes(ByVal strDateIn As String) As String

        Return strDateIn.Substring(0, 4) & "-" &
               strDateIn.Substring(4, 2) & "-" &
               strDateIn.Substring(6, 2)

    End Function

    Public Sub ClearUniqueNumbers()

        Dim cn As New SqlConnection(strConnection)
        Dim strSQL As String
        Dim cmd As SqlCommand

        strSQL = "TRUNCATE TABLE " & IT3sUniqueNoTable
        cmd = New SqlCommand(strSQL, cn)
        Try
            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()
        Catch ex As Exception
        End Try
    End Sub

    Public Sub ClearDatabase(SubmissionPeriod As String)

        Dim cn As New SqlConnection(strConnection)
        Dim strSQL As String
        Dim cmd As SqlCommand

        strSQL = "DELETE FROM " & IT3sHeaderTable & " WHERE it3sPeriod = '" & SubmissionPeriod & "' AND TestRun = '" & SARSForm.TestFile & "' " &
                 "DELETE FROM " & IT3sClientDataTable & " WHERE it3sPeriod = '" & SubmissionPeriod & "' AND TestRun = '" & SARSForm.TestFile & "' " &
                 "DELETE FROM " & IT3sExceptionTable & " WHERE it3sPeriod = '" & SubmissionPeriod & "' AND TestRun = '" & SARSForm.TestFile & "'"
        cmd = New SqlCommand(strSQL, cn)
        Try
            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()
        Catch ex As Exception
        End Try
    End Sub

    'Private Function ConsistsOfSpaces(ByVal strString As String) As Boolean

    '    Dim NonSpaceFound As Boolean = False
    '    Dim i As Integer

    '    i = 0
    '    While i > strString.Length - 1 And Not NonSpaceFound
    '        If strString.Substring(i, 1) <> " " Then
    '            NonSpaceFound = True
    '        End If
    '        i += 1
    '    End While
    '    Return Not NonSpaceFound
    'End Function

    Private Function TaxReferenceModulusCheck(ByVal strTaxRef As String) As Boolean
        '1.  Add all digits in even positions together (except last digit)
        '2.  Double the value of each digit in odd positions and add together
        '    (if a doubled value is greater than 10, add the digits, e.g. 9 * 2 = 18, so value is 1 + 8 = 9 )
        '3.  Add odd and even sums from steps 1 and 2
        '4.  If last digit of result is zero, check digit = 0, else check digit is 10 - last digit

        Dim i As Byte
        Dim Calc As Integer
        Dim Odds As Integer = 0
        Dim Evens As Integer = 0
        Dim OddsAndEvens As Integer
        Dim CheckDigit As String

        'Step 1:
        For i = 1 To lenTaxRef - 2 Step 2
            Evens = Evens + strTaxRef.Substring(i, 1)
        Next

        'Step 2:
        For i = 0 To lenTaxRef - 1 Step 2
            Calc = strTaxRef.Substring(i, 1) * 2
            If Calc > 9 Then
                Calc = 1 + Calc Mod 10
            End If
            Odds = Odds + Calc
        Next

        'Step 3:
        OddsAndEvens = Odds + Evens

        'Step 4:
        Select Case OddsAndEvens Mod 10
            Case 0
                CheckDigit = "0"
            Case Else
                CheckDigit = CStr(10 - (OddsAndEvens Mod 10))
        End Select

        'Step 5:
        If strTaxRef.Substring(lenTaxRef - 1, 1) <> CheckDigit Then
            Return False
        Else
            Return True
        End If
    End Function

    'Private Function TaxReferenceModulusCheck(ByVal strTaxRef As String) As Boolean
    '    '1.  Add all digits in odd positions together, except the last digit
    '    '2.  Double the value of each digit in even positions and add together.
    '    '    (if a doubled value is greater than 10, add the digits, e.g. 9 * 2 = 18, so value is 1 + 8 = 9 )
    '    '3.  Add odd and even sums from steps 1 and 2
    '    '4.  Check digit = 10 - sum

    '    Dim i As Byte
    '    Dim Calc As Integer
    '    Dim Odds As Integer = 0
    '    Dim Evens As Integer = 0
    '    Dim OddsAndEvens As Integer

    '    For i = 0 To lenTaxRef - 1 Step 2
    '        Calc = strTaxRef.Substring(i, 1) * 2
    '        If Calc > 10 Then
    '            Calc = 1 + Calc Mod 10
    '        End If
    '        Odds = Odds + Calc
    '    Next
    '    For i = 1 To lenTaxRef - 2 Step 2
    '        Evens = Evens + strTaxRef.Substring(i, 1)
    '    Next
    '    OddsAndEvens = Odds + Evens
    '    If 10 - OddsAndEvens Mod 10 <> strTaxRef.Substring(lenTaxRef - 1, 1) Then
    '        Return False
    '    Else
    '        Return True
    '    End If
    'End Function

    Private Function IDNoModulusCheck(ByVal strCIF As String, ByVal strIDNo As String) As String
        '1.  Add all digits in odd positions together, except the last digit
        '2.  Concatenate all digits in even positions and double the result
        '3.  Add all digits in result of step 2 above
        '4.  Add results from steps 1 and 3 together
        '5.  If last digit of result is zero, check digit = 0, else check digit is 10 - last digit

        Dim i As Byte
        Dim Odds As Integer = 0
        Dim EvenDigits As String = ""
        Dim Evens As Integer = 0
        Dim OddsAndEvens As Integer
        Dim CheckDigit As String

        If strIDNo.Length > 0 Then
            If IsNumeric(strIDNo) Then
                If CLng(strIDNo) > 0 Then
                    If strIDNo.Length = 13 Then

                        'Step 1:
                        For i = 0 To lenID - 2 Step 2
                            Odds = Odds + strIDNo.Substring(i, 1)
                        Next

                        'Step 2:
                        For i = 1 To lenID - 1 Step 2
                            EvenDigits = EvenDigits & strIDNo.Substring(i, 1)
                        Next
                        EvenDigits = CStr(CInt(EvenDigits) * 2)

                        'Step 3:
                        For i = 0 To EvenDigits.Length - 1
                            Evens = Evens + EvenDigits.Substring(i, 1)
                        Next

                        'Step 4:
                        OddsAndEvens = Odds + Evens

                        'Step 5:
                        Select Case OddsAndEvens Mod 10
                            Case 0
                                CheckDigit = "0"
                            Case Else
                                CheckDigit = CStr(10 - (OddsAndEvens Mod 10))
                        End Select

                        'Step 6:
                        If strIDNo.Substring(lenID - 1, 1) <> CheckDigit Then
                            If Not InvalidIDExceptionRaised Then
                                WriteExceptionToDatabase(IT3sExceptionTable, strCIF, "ID Number", "ID Number " & strIDNo & " does not pass modulus check")
                                InvalidIDExceptionRaised = True
                            End If
                            Return ""
                        Else
                            Return strIDNo
                        End If

                    Else
                        If Not InvalidIDExceptionRaised Then
                            WriteExceptionToDatabase(IT3sExceptionTable, strCIF, "ID Number", "ID Number " & strIDNo & " is incorrect length")
                            InvalidIDExceptionRaised = True
                        End If
                        Return ""
                    End If
                Else
                    If Not InvalidIDExceptionRaised Then
                        WriteExceptionToDatabase(IT3sExceptionTable, strCIF, "ID Number", "No ID Number found")
                        InvalidIDExceptionRaised = True
                    End If
                    Return ""
                End If
            Else
                If Not InvalidIDExceptionRaised Then
                    WriteExceptionToDatabase(IT3sExceptionTable, strCIF, "ID Number", "ID Number " & strIDNo & " is not numeric")
                    InvalidIDExceptionRaised = True
                End If
                Return ""
            End If
        Else
            If Not InvalidIDExceptionRaised Then
                WriteExceptionToDatabase(IT3sExceptionTable, strCIF, "ID Number", "No ID Number found")
                InvalidIDExceptionRaised = True
            End If
            Return ""
        End If
    End Function

    'Private Function IDNoModulusCheck(ByVal strCIF As String, ByVal strIDNo As String) As String
    '    '1.  Add all digits in even positions together, except the last digit
    '    '2.  Double the value of each digit in odd positions and add together.
    '    '    (if a doubled value is greater than 10, add the digits, e.g. 9 * 2 = 18, so value is 1 + 8 = 9 )
    '    '3.  Add odd and even sums from steps 1 and 2
    '    '4.  Check digit = 10 - sum

    '    Dim i As Byte
    '    Dim Calc As Integer
    '    Dim Odds As Integer = 0
    '    Dim Evens As Integer = 0
    '    Dim OddsAndEvens As Integer

    '    'If strIDNo.Length > 0 And Not ConsistsOfSpaces(strIDNo) Then
    '    If strIDNo.Length > 0 Then
    '        If IsNumeric(strIDNo) Then
    '            If CLng(strIDNo) > 0 Then
    '                For i = 1 To lenID - 1 Step 2
    '                    Calc = strIDNo.Substring(i, 1) * 2
    '                    If Calc > 10 Then
    '                        Calc = 1 + Calc Mod 10
    '                    End If
    '                    Evens = Evens + Calc
    '                Next
    '                For i = 0 To lenID - 2 Step 2
    '                    Odds = Odds + strIDNo.Substring(i, 1)
    '                Next
    '                OddsAndEvens = Odds + Evens
    '                If 10 - OddsAndEvens Mod 10 <> strIDNo.Substring(lenID - 1, 1) Then
    '                    If Not InvalidIDExceptionRaised Then
    '                        WriteExceptionToDatabase(IT3sExceptionTable, strCIF, "ID Number", "ID Number " & strIDNo & " does not pass modulus check")
    '                        InvalidIDExceptionRaised = True
    '                    End If
    '                    Return ""
    '                Else
    '                    Return strIDNo
    '                End If
    '            Else
    '                If Not InvalidIDExceptionRaised Then
    '                    WriteExceptionToDatabase(IT3sExceptionTable, strCIF, "ID Number", "No ID Number found")
    '                    InvalidIDExceptionRaised = True
    '                End If
    '                Return ""
    '            End If
    '        Else
    '            If Not InvalidIDExceptionRaised Then
    '                WriteExceptionToDatabase(IT3sExceptionTable, strCIF, "ID Number", "ID Number " & strIDNo & " is not numeric")
    '                InvalidIDExceptionRaised = True
    '            End If
    '            Return ""
    '        End If
    '    Else
    '        If Not InvalidIDExceptionRaised Then
    '            WriteExceptionToDatabase(IT3sExceptionTable, strCIF, "ID Number", "No ID Number found")
    '            InvalidIDExceptionRaised = True
    '        End If
    '        Return ""
    '    End If
    'End Function

    Private Function ValidTaxReference(ByVal strCIF As String, ByVal strTaxRef As String) As String

        If strTaxRef.Length > 0 Then
            If strTaxRef <> "0000000000" And Not strTaxRef.ToUpper.Contains("NA") Then
                If IsNumeric(strTaxRef) Then
                    Select Case strTaxRef.Substring(0, 1)
                        Case "0", "1", "2", "3", "9"
                            If TaxReferenceModulusCheck(strTaxRef) Then
                                ValidTaxReference = strTaxRef
                            Else
                                If Not InvalidTaxNoExceptionRaised Then
                                    WriteExceptionToDatabase(IT3sExceptionTable, strCIF, "Tax Reference", "Tax Reference " & strTaxRef & " does not pass modulus check")
                                    InvalidTaxNoExceptionRaised = True
                                End If
                                ValidTaxReference = ""
                            End If
                        Case Else
                            ValidTaxReference = ""
                    End Select
                Else
                    If Not InvalidTaxNoExceptionRaised Then
                        WriteExceptionToDatabase(IT3sExceptionTable, strCIF, "Tax Reference", "Non-numeric Tax Reference " & strTaxRef)
                        InvalidTaxNoExceptionRaised = True
                    End If
                    ValidTaxReference = ""
                End If
            Else
                ValidTaxReference = ""
            End If
        Else
            ValidTaxReference = ""
        End If
    End Function

    'Private Function ValidCompanyRegNo(ByVal strCIF As String, ByVal strRegNo As String) As String

    '    Dim IsValid As Boolean = True

    '    'If strRegNo.Length <> 12 Then
    '    If strRegNo.Length < 14 Then
    '        IsValid = False
    '    End If
    '    If IsValid Then
    '        If Not (strRegNo.Substring(4, 1) = "/" And strRegNo.Substring(11, 1) = "/") Then
    '            IsValid = False
    '        End If
    '    End If
    '    If IsValid Then
    '        If Not IsNumeric(strRegNo.Substring(0, 4)) Then
    '            IsValid = False
    '        End If
    '    End If
    '    If IsValid Then
    '        If Not IsNumeric(strRegNo.Substring(5, 6)) Then
    '            IsValid = False
    '        End If
    '    End If
    '    If IsValid Then
    '        If Not IsNumeric(strRegNo.Substring(12, 2)) Then
    '            IsValid = False
    '        End If
    '    End If
    '    If IsValid Then
    '        If Not (strRegNo.Substring(0, 4) >= 1800 And strRegNo.Substring(0, 4) <= SARSForm.PeriodEnd.Substring(0, 4)) Then
    '            IsValid = False
    '        End If
    '    End If
    '    If IsValid Then
    '        Select Case strRegNo.Substring(strRegNo.Length - 2) Mod 100
    '            Case 6 To 11
    '            Case 20 To 26
    '            Case Else
    '                IsValid = False
    '        End Select
    '    End If
    '    If Not IsValid Then
    '        If Not InvalidCoRegNoExceptionRaised Then
    '            WriteExceptionToDatabase(IT3bExceptionTable, strCIF, "Registration Number", "Company/CC registration number " & strRegNo & " invalid")
    '            InvalidCoRegNoExceptionRaised = True
    '        End If
    '        Return ""
    '    Else
    '        Return strRegNo
    '    End If
    'End Function

    Private Function ValidCompanyRegNo(ByVal strCIF As String, ByVal strRegNo As String) As String

        Dim IsValid As Boolean = True

        strRegNo = strRegNo.Replace(" ", "")

        If strRegNo.Length <> 14 Then
            IsValid = False
        End If

        If IsValid AndAlso Not IsNumeric(strRegNo.Replace("/", "")) Then
            IsValid = False
        End If

        If IsValid Then

            Dim col() As String = strRegNo.Split("/")

            If col.GetUpperBound(0) <> 2 Then
                IsValid = False
            End If

            If IsValid AndAlso Not (col(0).Length = 4 And col(1).Length = 6 And col(2).Length = 2) Then
                IsValid = False
            End If

            If IsValid AndAlso (CInt(col(0)) < 1800 Or CInt(col(0)) > Year(Now)) Then
                IsValid = False
            End If

            If IsValid Then
                Dim fnd As Boolean = False
                Dim lst() As Integer = {6, 7, 8, 9, 10, 11, 20, 21, 22, 23, 24, 25, 26, 30, 31}
                For Each i As Integer In lst
                    If i = CInt(col(2)) Then
                        fnd = True
                    End If
                Next
                IsValid = fnd
            End If

        End If

        If IsValid Then
            Return strRegNo
        Else
            If Not InvalidCoRegNoExceptionRaised Then
                WriteExceptionToDatabase(IT3sExceptionTable, strCIF, "Registration Number", "Company/CC registration number " & strRegNo & " invalid")
                InvalidCoRegNoExceptionRaised = True
            End If
            Return ""
        End If

    End Function

    Private Function ValidTrustRegNo(ByVal strCIF As String, ByVal strRegNo As String) As String

        Dim IsValid As Boolean = True

        If strRegNo.Length < 8 Then
            IsValid = False
        End If

        If IsValid Then
            Return strRegNo
        Else
            If Not InvalidTrustRegNoExceptionRaised Then
                WriteExceptionToDatabase(IT3sExceptionTable, strCIF, "Registration Number", "Trust registration number " & strRegNo & " invalid")
                InvalidTrustRegNoExceptionRaised = True
            End If
            Return ""
        End If

    End Function

    'Private Function ValidTrustRegNo(ByVal strCIF As String, ByVal strRegNo As String) As String

    '    Dim IsValid As Boolean = True

    '    'If strRegNo.Length <> 8 Then
    '    If strRegNo.Length < 8 Then
    '        IsValid = False
    '    ElseIf Not IsNumeric(strRegNo.Substring(strRegNo.Length - 4)) Then
    '        IsValid = False
    '    End If
    '    If IsValid Then
    '        If Not (strRegNo.Substring(strRegNo.Length - 4) >= 1900 And strRegNo.Substring(strRegNo.Length - 4) <= SARSForm.PeriodEnd.Substring(0, 4)) Then
    '            IsValid = False
    '        End If
    '    End If
    '    If Not IsValid Then
    '        If Not InvalidTrustRegNoExceptionRaised Then
    '            WriteExceptionToDatabase(IT3bExceptionTable, strCIF, "Registration Number", "Trust registration number " & strRegNo & " invalid")
    '            InvalidTrustRegNoExceptionRaised = True
    '        End If
    '        Return ""
    '    Else
    '        Return strRegNo
    '    End If
    'End Function

    Private Function ValidateDOB(ByVal strCIF As String, ByVal strDOB As String, ByVal strIDNo As String) As String

        Dim tmpDOB As String

        If strDOB.Substring(2) = strIDNo.Substring(0, 6) Then
            Return strDOB
        Else
            tmpDOB = strIDNo.Substring(0, 6)
            If tmpDOB.Substring(0, 2) < Now.Year.ToString.Substring(2) Then
                tmpDOB = "20" & tmpDOB
            Else
                tmpDOB = "19" & tmpDOB
            End If
            WriteExceptionToDatabase(IT3sExceptionTable, strCIF, "Date of Birth", "Date of Birth " & strDOB & " does not correspond with ID No. " & strIDNo & ".  Assuming DOB to be " & tmpDOB)
            Return tmpDOB
        End If
    End Function

    Private Function GetAddress(ByVal cifId As Integer, ByVal strCIF As String, ByVal AddressType As String) As Array

        Dim Address(8) As String
        Dim cmd As SqlCommand
        Dim strSQL As String
        If AddressType = "Postal" Then
            strSQL = "SELECT TOP 1 tblAddress.addrId, tblAddress.adtid, tblAddress.cifId, tblAddress.ctyId, " &
                            "tblAddress.addrLine1, tblAddress.addrLine2, tblAddress.addrLine3, tblAddress.addrLine4, " &
                            "tblAddress.addrSuburb, tblAddress.cityId AS addrCityId, tblAddress.addrProvince, " &
                            "tblAddress.addrCode, tblAddress.Deleted AS AddressDeleted, mblCity.cityId, mblCity.cityDesc, " &
                            "mblCity.cityPolicyCode, mblCity.cityOrder, mblCity.Deleted AS CityDeleted " &
                     "FROM   tblAddress LEFT OUTER JOIN " &
                            "mblCity ON tblAddress.cityId = mblCity.cityId " &
                     "WHERE tblAddress.cifId = " & cifId & " AND tblAddress.adtid = 2 " &
                                                            "AND tblAddress.Deleted = 0 " &
                     "ORDER BY tblAddress.addrId DESC"
        Else
            strSQL = "SELECT TOP 1 tblAddress.addrId, tblAddress.adtid, tblAddress.cifId, tblAddress.ctyId, " &
                            "tblAddress.addrLine1, tblAddress.addrLine2, tblAddress.addrLine3, tblAddress.addrLine4, " &
                            "tblAddress.addrSuburb, tblAddress.cityId AS addrCityId, tblAddress.addrProvince, " &
                            "tblAddress.addrCode, tblAddress.Deleted AS AddressDeleted, mblCity.cityId, mblCity.cityDesc, " &
                            "mblCity.cityPolicyCode, mblCity.cityOrder, mblCity.Deleted AS CityDeleted " &
                     "FROM   tblAddress LEFT OUTER JOIN " &
                            "mblCity ON tblAddress.cityId = mblCity.cityId " &
                     "WHERE tblAddress.cifId = " & cifId & " AND " &
                           "(tblAddress.adtid = 1 OR tblAddress.adtid = 14 OR tblAddress.adtid = 16) " &
                                                            "AND tblAddress.Deleted = 0 " &
                     "ORDER BY tblAddress.addrId DESC"
        End If

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader

        Address(0) = "No Address"
        cmd = New SqlCommand(strSQL, cn)
        Try
            cn.Open()
            drSQL = cmd.ExecuteReader()
            If drSQL.Read() Then
                Address(1) = FixNull(drSQL.Item("addrLine1")).Replace("|", "/")
                Address(2) = FixNull(drSQL.Item("addrLine2")).Replace("|", "/")
                Address(3) = FixNull(drSQL.Item("addrLine3")).Replace("|", "/")
                Address(4) = FixNull(drSQL.Item("addrLine4")).Replace("|", "/")
                Address(5) = FixNull(drSQL.Item("addrSuburb"))
                Address(6) = FixNull(drSQL.Item("cityDesc"))
                Address(7) = DropChar(FixNull(drSQL.Item("addrCode")), " ")
                Address(8) = FixNull(drSQL.Item("addrProvince"))
                Address(7) = Address(7).Replace("-", " ").Replace(",", " ")
                While Address(7).Length > 10
                    If Address(7).Contains(" ") Then
                        Address(7) = Address(7).Replace(" ", "")
                    Else
                        Address(7) = Address(7).Substring(0, 10)
                    End If
                End While
                If Address(1) = "" And Address(2) = "" And Address(3) = "" And Address(4) = "" Then
                    WriteExceptionToDatabase(IT3sExceptionTable, strCIF, "Address", "Address is blank")
                ElseIf Address(1).ToUpper.Contains("UNKNOWN") Or Address(2).ToUpper.Contains("UNKNOWN") Or Address(3).ToUpper.Contains("UNKNOWN") Or Address(4).ToUpper.Contains("UNKNOWN") Then
                    WriteExceptionToDatabase(IT3sExceptionTable, strCIF, "Address", "Address is flagged as Unknown")
                ElseIf IsDBNull(drSQL.Item("addrCityId")) Then
                    WriteExceptionToDatabase(IT3sExceptionTable, strCIF, "Address", "Address has no city, or city incorrectly encoded")
                    Address(0) = "Address Found"
                ElseIf drSQL.Item("CityDeleted") Then
                    WriteExceptionToDatabase(IT3sExceptionTable, strCIF, "Address", "Address is linked to a deleted city")
                    Address(0) = "Address Found"
                Else
                    Address(0) = "Address Found"
                End If
            End If
            cn.Close()
        Catch ex As Exception
            WriteExceptionToDatabase(IT3sExceptionTable, strCIF, "Address", "An error occurred while trying to obtain client address")
            Address(0) = "No Address"
        End Try
        If Address(0) = "Address Found" Then
            If Address(1).Length > 35 Then
                Address(1) = Address(1).Substring(0, 35)
            End If
            If Address(2).Length > 35 Then
                Address(2) = Address(2).Substring(0, 35)
            End If
            If Address(3).Length > 35 Then
                Address(3) = Address(3).Substring(0, 35)
            End If
            If Address(4).Length > 35 Then
                Address(4) = Address(4).Substring(0, 35)
            End If
        End If
        Return Address
    End Function

    Private Function GetContactDetails(ByVal cifId As Integer, ByVal strCIF As String) As Array

        Dim Contacts() As String = {"", "", ""}
        Dim cmd As SqlCommand
        Dim strSQL As String

        strSQL = "SELECT	 ConsolidatedContacts.cifId 
		                    ,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(ConsolidatedContacts.TelephoneNo, '-', ''), ' ', ''), '(', ''), ')', ''), '+', '00') AS TelephoneNo
		                    ,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(ConsolidatedContacts.CellPhoneNo, '-', ''), ' ', ''), '(', ''), ')', ''), '+', '00') AS CellphoneNo
		                    ,CASE WHEN COALESCE(tblEmailAddress.emIsValid, 1) = 1 THEN
			                    ConsolidatedContacts.EMailAddress
		                     ELSE
			                    ''
		                     END AS EmailAddress
                    FROM   (SELECT	 cifId
				                    ,CASE WHEN Tel1 <> '' THEN
					                    Tel1
				                     ELSE
				                     CASE WHEN Tel2 <> '' THEN
					                    Tel2
				                     ELSE
					                    Tel3
				                     END END AS TelephoneNo
				                    ,Cell AS CellPhoneNo
				                    ,CASE WHEN EMail <> '' THEN
					                    EMail
				                     ELSE
					                    EmailAlt
				                     END AS EMailAddress
		                    FROM   (SELECT	 tblCIF.cifId
						                    ,COALESCE(tblCIFIndividual.cifiTelHome, tblCIFCompany.cifcTel1, '') AS Tel1
						                    ,COALESCE(tblCIFIndividual.cifiTelWork, tblCIFCompany.cifcTel2, '') AS Tel2
						                    ,COALESCE(tblCIFIndividual.cifiTelAlternate, tblCIFCompany.cifcTel3, '') AS Tel3
						                    ,COALESCE(tblCIFIndividual.cifiCell, tblCIFCompany.cifcCell, '') AS Cell
						                    ,COALESCE(tblCIFIndividual.cifiEmail, tblCIFCompany.cifcEmail, '') AS EMail
						                    ,COALESCE(tblCIFIndividual.cifiAltEmail, tblCIFCompany.cifcAltEmail, '') AS EmailAlt
				                    FROM	tblCIF LEFT OUTER JOIN
						                    tblCIFCompany ON tblCIF.cifId = tblCIFCompany.cifId LEFT OUTER JOIN
						                    tblCIFIndividual ON tblCIF.cifId = tblCIFIndividual.cifId
				                    WHERE	tblCIF.cifId = " & cifId & ") AS Contacts) AS ConsolidatedContacts LEFT OUTER JOIN
				                    tblEmailAddress ON ConsolidatedContacts.EMailAddress = tblEmailAddress.emAddress"

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader

        cmd = New SqlCommand(strSQL, cn)
        Try
            cn.Open()
            drSQL = cmd.ExecuteReader()
            If drSQL.Read() Then
                Contacts(0) = drSQL.Item("TelephoneNo")
                Contacts(1) = drSQL.Item("CellPhoneNo")
                Contacts(2) = drSQL.Item("EMailAddress")
            End If
            cn.Close()
        Catch ex As Exception
            WriteExceptionToDatabase(IT3sExceptionTable, strCIF, "Contact Details", "An error occurred while trying to obtain client contact details")
        End Try

        If Contacts(0).Length > 10 Then
            Contacts(0) = Contacts(0).Substring(0, 10)
        End If

        If Contacts(1).Length > 10 Then
            Contacts(1) = Contacts(1).Substring(0, 10)
        End If

        If Not IsNumeric(Contacts(0)) Then
            Contacts(0) = ""
        End If

        If Not IsNumeric(Contacts(1)) Then
            Contacts(1) = ""
        End If

        Return Contacts

    End Function

    'Public Function Old_FileHeader() As String

    '    Const SEC_ID As String = "H"
    '    Const INFO_TYPE As String = "IT3EXTRS"
    '    Const INFO_SUBTYPE As String = " "
    '    Const FILE_SERIES_CTL As String = "S"
    '    Const EXT_SYS As String = "GBSMUTUA"
    '    Const VER_NO As String = "1"

    '    Dim GEN_TIME As String
    '    Dim Now As DateTime = DateTime.Now

    '    GEN_TIME = Now.Year & FixedLengthString(Now.Month, 2, "Right", "0") & FixedLengthString(Now.Day, 2, "Right", "0") & _
    '               FixedLengthString(Now.Hour, 2, "Right", "0") & FixedLengthString(Now.Minute, 2, "Right", "0") & FixedLengthString(Now.Second, 2, "Right", "0")

    '    FileHeader = SEC_ID & INFO_TYPE & FixedLengthString(INFO_SUBTYPE, 8, "Left", " ") & _
    '                 SARSForm.TestFile & FILE_SERIES_CTL & EXT_SYS & FixedLengthString(VER_NO, 8, "Right", " ") & _
    '                 FixedLengthString(SARSForm.Reference, 14, "Right", " ") & GEN_TIME
    'End Function

    Public Function FileHeader(ByVal _5_UNIQUE_ID As String) As String

        Const _1_SEC_ID As String = "H"
        Const _2_HDR_TYPE As String = "GH"
        Dim _3_GEN_TIME As String
        Dim _4_LAYOUT_VERSION As String = FileLayoutVersion
        Const _6_SARS_REF As String = ""
        Dim _7_TEST_IND As String = SARSForm.txtTest.Text
        Dim _8_DATA_TYPE As String = FileDataType
        Const _9_CHANNEL_ID As String = "HTTPS"
        Dim _10_SOURCE_ID As String = FixedLengthString(SARSSourceID, 144, "Right", "-")
        Dim _11_GROUP_ID As String = SARSForm.txtReference.Text
        Const _12_GROUP_TOTAL As String = "GROUP_TOTAL"
        Dim _13_GROUP_ITEM As String = FileSeq
        Dim _14_SOURCE_SYSTEM As String = SourceSystem
        Dim _15_SOURCE_SYSTEM_VERSION As String = SourceSystemVersion
        Dim _16_CONTACT_PERSON_NAME As String = ContactPersonName
        Dim _17_CONTACT_PERSON_SURNAME As String = ContactPersonSurname
        Dim _18_TELEPHONE_1 As String = Telephone1
        Dim _19_TELEPHONE_2 As String = Telephone2
        Dim _20_CELL As String = CellNo
        Dim _21_EMAIL As String = EMail

        Dim TimeNow As DateTime = DateTime.Now

        _3_GEN_TIME = TimeNow.Year & "-" &
                      FixedLengthString(TimeNow.Month, 2, "Right", "0") & "-" &
                      FixedLengthString(TimeNow.Day, 2, "Right", "0") & "T" &
                      FixedLengthString(TimeNow.Hour, 2, "Right", "0") & ":" &
                      FixedLengthString(TimeNow.Minute, 2, "Right", "0") & ":" &
                      FixedLengthString(TimeNow.Second, 2, "Right", "0")

        FileHeader = _1_SEC_ID & DataDelimiter &
                     _2_HDR_TYPE & DataDelimiter &
                     _3_GEN_TIME & DataDelimiter &
                     _4_LAYOUT_VERSION & DataDelimiter &
                     _5_UNIQUE_ID & DataDelimiter &
                     _6_SARS_REF & DataDelimiter &
                     _7_TEST_IND & DataDelimiter &
                     _8_DATA_TYPE & DataDelimiter &
                     _9_CHANNEL_ID & DataDelimiter &
                     _10_SOURCE_ID & DataDelimiter &
                     _11_GROUP_ID & DataDelimiter &
                     _12_GROUP_TOTAL & DataDelimiter &
                     _13_GROUP_ITEM & DataDelimiter &
                     _14_SOURCE_SYSTEM & DataDelimiter &
                     _15_SOURCE_SYSTEM_VERSION & DataDelimiter &
                     _16_CONTACT_PERSON_NAME & DataDelimiter &
                     _17_CONTACT_PERSON_SURNAME & DataDelimiter &
                     _18_TELEPHONE_1 & DataDelimiter &
                     _19_TELEPHONE_2 & DataDelimiter &
                     _20_CELL & DataDelimiter &
                     _21_EMAIL
    End Function

    Private Function SubmittingEntityDetails() As String

        Const _22_SEC_ID As String = "H"
        Const _23_RECORD_TYPE As String = "SE"
        'Const _24_RECORD_STATUS As String = "A"
        'Dim _25_UNIQUE_NO As String = TotalRecords.ToString
        'Dim _26_ROW_NO As String = RecSeq.ToString
        Dim _24_TAX_YEAR As String
        Dim _25_PERIOD_START As String = DateWithDashes(SARSForm.PeriodStart(SARSForm.txtPeriodLength.Text))
        Dim _26_PERIOD_END As String = DateWithDashes(SARSForm.PeriodEnd)
        Dim _135_NATURE_OF_PERSON As String = NatureOfPerson
        Dim _27_REGISTERED_NAME As String = Institution
        Dim _136_TRADING_NAME As String = Institution
        Dim _137_REGISTRATION_NO As String = ""
        Dim _601_REGULATOR_REGISTRATION_NO As String = RegulatorRegistrationNo
        Dim _602_REGULATOR_DESIGNATION As String = RegulatorDesignation
        Dim _28_TAX_REFERENCE As String = InstitutionTaxReference
        Dim _29_BRANCH_CODE As String = BranchCode
        Dim _138_POSTAL_ADDRESS_1 As String = PostalAddress1
        Dim _139_POSTAL_ADDRESS_2 As String = PostalAddress2
        Dim _140_POSTAL_ADDRESS_3 As String = PostalAddress3
        Dim _141_POSTAL_ADDRESS_4 As String = PostalAddress4
        Dim _142_POSTAL_CODE As String = PostalCode

        If SARSForm.PeriodEnd.Substring(4, 2) > 2 Then
            _24_TAX_YEAR = SARSForm.PeriodEnd.Substring(0, 4) + 1
        Else
            _24_TAX_YEAR = SARSForm.PeriodEnd.Substring(0, 4)
        End If

        SubmittingEntityDetails = _22_SEC_ID & DataDelimiter &
                                  _23_RECORD_TYPE & DataDelimiter &
                                  _24_TAX_YEAR & DataDelimiter &
                                  _25_PERIOD_START & DataDelimiter &
                                  _26_PERIOD_END & DataDelimiter &
                                  _135_NATURE_OF_PERSON & DataDelimiter &
                                  _27_REGISTERED_NAME & DataDelimiter &
                                  _136_TRADING_NAME & DataDelimiter &
                                  _137_REGISTRATION_NO & DataDelimiter &
                                  _601_REGULATOR_REGISTRATION_NO & DataDelimiter &
                                  _602_REGULATOR_DESIGNATION & DataDelimiter &
                                  _28_TAX_REFERENCE & DataDelimiter &
                                  _29_BRANCH_CODE & DataDelimiter &
                                  _138_POSTAL_ADDRESS_1 & DataDelimiter &
                                  _139_POSTAL_ADDRESS_2 & DataDelimiter &
                                  _140_POSTAL_ADDRESS_3 & DataDelimiter &
                                  _141_POSTAL_ADDRESS_4 & DataDelimiter &
                                  _142_POSTAL_CODE
    End Function

    Private Function GetExistingUniqueNo(cifNo As String, recType As String, TestFlag As String, Optional accNO As String = "", Optional SourceCode As String = "") As String

        Dim strSQL As String
        Dim it3sRecord As String = ""

        If accNO = "" Then
            strSQL = "SELECT it3sRecord " &
                     "FROM " & IT3sClientDataTable & " " &
                     "WHERE  recId = (SELECT MIN(recId) AS recId " &
                                     "FROM " & IT3sClientDataTable & " " &
                                     "WHERE  it3sRecord LIKE '%|" & cifNo & "|%' " &
                                                "AND it3sRecord LIKE '%|" & recType & "|%' " &
                                                "AND it3sPeriod = '" & it3sPeriod & "' " &
                                                "AND TestRun = '" & TestFlag & "')"
        Else
            If SourceCode = "" Then
                strSQL = "SELECT it3sRecord " &
                         "FROM " & IT3sClientDataTable & " " &
                         "WHERE  recId = (SELECT MIN(recId) AS recId " &
                                         "FROM " & IT3sClientDataTable & " " &
                                         "WHERE  it3sRecord LIKE '%|" & cifNo & "|%' " &
                                                    "AND it3sRecord LIKE '%|" & accNO & "|%' " &
                                                    "AND it3sRecord LIKE '%|" & recType & "|%' " &
                                                    "AND it3sPeriod = '" & it3sPeriod & "' " &
                                                    "AND TestRun = '" & TestFlag & "')"
            Else
                strSQL = "SELECT it3sRecord " &
                         "FROM " & IT3sClientDataTable & " " &
                         "WHERE  recId = (SELECT MIN(recId) AS recId " &
                                         "FROM " & IT3sClientDataTable & " " &
                                         "WHERE  it3sRecord LIKE '%|" & cifNo & "|%' " &
                                                    "AND it3sRecord LIKE '%|" & recType & "|%' " &
                                                    "AND it3sPeriod = '" & it3sPeriod & "' " &
                                                    "AND it3sRecord LIKE '%|" & SourceCode & "|%' " &
                                                    "AND TestRun = '" & TestFlag & "')"
            End If
        End If

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)

        cn.Open()
        drSQL = cmd.ExecuteReader()
        Try
            drSQL.Read()
            it3sRecord = drSQL.Item(0)
        Catch ex As Exception
        End Try
        cn.Close()

        If it3sRecord <> "" Then
            Return GetColumnFromRecord(4, it3sRecord, DataDelimiter)
        Else
            Return "Unique_No"
        End If

    End Function

    'Private Function GetExistingUniqueNo(cifNo As String, recType As String, TestFlag As String, Optional accNO As String = "") As String

    '    Dim strSQL As String
    '    Dim it3sRecord As String = ""

    '    If accNO = "" Then
    '        strSQL = "SELECT it3sRecord " &
    '                 "FROM " & IT3sClientDataTable & " " &
    '                 "WHERE  recId = (SELECT MIN(recId) AS recId " &
    '                                 "FROM " & IT3sClientDataTable & " " &
    '                                 "WHERE  it3sRecord LIKE '%|" & cifNo & "|%' " &
    '                                            "AND it3sRecord LIKE '%|" & recType & "|%' " &
    '                                            "AND it3sPeriod = '" & it3sPeriod & "' " &
    '                                            "AND TestRun = '" & TestFlag & "')"
    '    Else
    '        strSQL = "SELECT it3sRecord " &
    '                 "FROM " & IT3sClientDataTable & " " &
    '                 "WHERE  recId = (SELECT MIN(recId) AS recId " &
    '                                 "FROM " & IT3sClientDataTable & " " &
    '                                 "WHERE  it3sRecord LIKE '%|" & cifNo & "|%' " &
    '                                            "AND it3sRecord LIKE '%|" & accNO & "|%' " &
    '                                            "AND it3sRecord LIKE '%|" & recType & "|%' " &
    '                                            "AND it3sPeriod = '" & it3sPeriod & "' " &
    '                                            "AND TestRun = '" & TestFlag & "')"
    '    End If

    '    Dim cn As SqlConnection = New SqlConnection(strConnection)
    '    Dim drSQL As SqlDataReader
    '    Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)

    '    cn.Open()
    '    drSQL = cmd.ExecuteReader()
    '    Try
    '        drSQL.Read()
    '        it3sRecord = drSQL.Item(0)
    '    Catch ex As Exception
    '    End Try
    '    cn.Close()

    '    If it3sRecord <> "" Then
    '        Return GetColumnFromRecord(4, it3sRecord, DataDelimiter)
    '    Else
    '        Return "Unique_No"
    '    End If

    'End Function

    'Private Function GetAHFDUniqueNo(strTrnId As String, TestFlag As String) As String

    '    Dim strSQL As String
    '    Dim it3sRecord As String = ""

    '    strSQL = "SELECT it3sRecord " &
    '             "FROM " & IT3sClientDataTable & " " &
    '             "WHERE  recId = (SELECT MIN(recId) AS recId " &
    '                            "FROM " & IT3sClientDataTable & " " &
    '                            "WHERE  it3sRecord LIKE '%|" & GetAccountNo(strTrnId) & "|%' " &
    '                                    "AND it3sRecord LIKE '%|AHFD|%' " &
    '                                    "AND it3sPeriod = '" & it3sPeriod & "' " &
    '                                    "AND TestRun = '" & TestFlag & "')"

    '    Dim cn As SqlConnection = New SqlConnection(strConnection)
    '    Dim drSQL As SqlDataReader
    '    Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)

    '    cn.Open()
    '    drSQL = cmd.ExecuteReader()
    '    Try
    '        drSQL.Read()
    '        it3sRecord = drSQL.Item(0)
    '    Catch ex As Exception
    '    End Try
    '    cn.Close()

    '    If it3sRecord <> "" Then
    '        Return GetColumnFromRecord(4, it3sRecord, DataDelimiter)
    '    Else
    '        Return "AHFD_Unique_No"
    '    End If

    'End Function

    Private Function GetAHFDUniqueNo(strTrnId As String, TestFlag As String) As String

        Dim strSQL As String
        Dim it3sRecId As Integer = 0
        Dim AccountNo As String

        If strTrnId.Contains("<") Then
            AccountNo = strTrnId.Replace("<", "").Replace(">", "")
        Else
            AccountNo = GetAccountNo(strTrnId)
        End If

        strSQL = "SELECT MAX(recId) AS recId " &
                    "FROM " & IT3sClientDataTable & " " &
                    "WHERE  it3sRecord LIKE '%|" & AccountNo & "|%' " &
                            "AND it3sRecord LIKE '%|AHFD|%' " &
                            "AND it3sPeriod = '" & it3sPeriod & "' " &
                            "AND TestRun = '" & TestFlag & "'"

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)

        cn.Open()
        drSQL = cmd.ExecuteReader()
        Try
            drSQL.Read()
            it3sRecId = drSQL.Item(0)
        Catch ex As Exception
        End Try
        cn.Close()

        If it3sRecId <> 0 Then
            Return GetUniqueNumber(it3sRecId)
        Else
            Return "AHFD_Unique_No"
        End If

    End Function

    Private Function GetAccountNo(trnId As String) As String

        Dim strSQL As String
        Dim strAccountNo As String = ""

        strSQL = "SELECT	tblAccount.accNO
                    FROM	tblTransactions INNER JOIN
		                    tblAccount ON tblTransactions.accId = tblAccount.accId
                    WHERE	tblTransactions.trnId = " & trnId

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)

        cn.Open()
        drSQL = cmd.ExecuteReader()
        Try
            drSQL.Read()
            strAccountNo = drSQL.Item(0)
        Catch ex As Exception
        End Try
        cn.Close()

        Return strAccountNo

    End Function

    Private Function RecordCompletelyRejected(recNo As String, TestFlag As String) As Boolean

        Dim strSQL As String
        Dim IT3sResponse As String = ""
        Dim EntireRecordRejected As Boolean = False

        strSQL = "SELECT IT3sResponse " &
                 "FROM " & IT3sResponseTable & " " &
                 "WHERE   IT3sResponse LIKE '%|" & recNo & "|%' " &
                            "AND IT3sPeriod = '" & it3sPeriod & "' " &
                            "AND IT3sSubmissionNo = " & it3sSubmissionNo - 1 & " " &
                            "AND TestRun = '" & TestFlag & "'"

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)

        Try
            cn.Open()
            drSQL = cmd.ExecuteReader()
            While drSQL.Read()

                IT3sResponse = drSQL.Item(0)

                If IT3sResponse <> "" Then
                    If GetColumnFromRecord(5, IT3sResponse, DataDelimiter) = "R" Then
                        EntireRecordRejected = True
                    End If
                End If

            End While
        Catch ex As Exception
        End Try
        cn.Close()

        Return EntireRecordRejected

    End Function

    'Private Function RecordCompletelyRejected(recNo As String, TestFlag As String) As Boolean

    '    Dim strSQL As String
    '    Dim it3bResponse As String = ""
    '    Dim EntireRecordRejected As Boolean = False
    '    Dim ScopeOfRejection As String = ""

    '    strSQL = "SELECT it3bResponse " &
    '             "FROM " & IT3bResponseTable & " " &
    '             "WHERE   it3bResponse LIKE '%|" & recNo & "|%' " &
    '                        "AND it3bPeriod = '" & it3bPeriod & "' " &
    '                        "AND it3bSubmissionNo = " & it3bSubmissionNo - 1 & " " &
    '                        "AND TestRun = '" & TestFlag & "'"

    '    Dim cn As SqlConnection = New SqlConnection(strConnection)
    '    Dim drSQL As SqlDataReader
    '    Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)

    '    Try
    '        cn.Open()
    '        drSQL = cmd.ExecuteReader()
    '        While drSQL.Read()

    '            it3bResponse = drSQL.Item(0)

    '            If it3bResponse <> "" Then

    '                If GetColumnFromRecord(5, it3bResponse, DataDelimiter) = "R" Then
    '                    ScopeOfRejection = GetColumnFromRecord(4, it3bResponse, DataDelimiter)
    '                Else
    '                    ScopeOfRejection = ""
    '                End If

    '                If ScopeOfRejection.ToUpper = "ENTIRE RECORD" Or (ScopeOfRejection.Contains("96.") And ScopeOfRejection.ToUpper.Contains("UNIQUE NUMBER")) Then
    '                    EntireRecordRejected = True
    '                End If

    '            End If

    '        End While
    '    Catch ex As Exception
    '    End Try
    '    cn.Close()

    '    Return EntireRecordRejected

    'End Function

    'Private Function RecordCompletelyRejected(recNo As String, recType As String, TestFlag As String) As Boolean

    '    Dim strSQL As String
    '    Dim it3bResponse As String = ""
    '    Dim recStatus As String = ""
    '    Dim ScopeOfRejection As String = ""

    '    strSQL = "SELECT it3bResponse " &
    '             "FROM " & IT3bResponseTable & " " &
    '             "WHERE   it3bResponse LIKE '%|" & recNo & "|%' " &
    '                        "AND it3bPeriod = '" & it3bPeriod & "' " &
    '                        "AND it3bSubmissionNo = " & it3bSubmissionNo - 1 & " " &
    '                        "AND TestRun = '" & TestFlag & "'"

    '    Dim cn As SqlConnection = New SqlConnection(strConnection)
    '    Dim drSQL As SqlDataReader
    '    Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)

    '    cn.Open()
    '    drSQL = cmd.ExecuteReader()
    '    Try
    '        drSQL.Read()
    '        it3bResponse = drSQL.Item(0)
    '    Catch ex As Exception
    '    End Try
    '    cn.Close()

    '    If it3bResponse <> "" Then
    '        recStatus = GetColumnFromRecord(5, it3bResponse, DataDelimiter)
    '        ScopeOfRejection = GetColumnFromRecord(4, it3bResponse, DataDelimiter)
    '    End If

    '    If recStatus = "R" And ScopeOfRejection = "Entire Record" Then
    '        Return True
    '    Else
    '        Return False
    '    End If

    'End Function

    Private Function PersonalDetailsRecord(ByVal CIFInfo As Object, ByRef ClientHasReferences As Boolean) As String

        Const _30_SEC_ID As String = "B"
        Const _31_RECORD_TYPE As String = "AHDD"
        Dim _32_RECORD_STATUS As String
        Dim _33_UNIQUE_NO As String = "Unique_No"     'TotalRecords.ToString
        Dim _34_ROW_NO As String = "Rec_No"
        Dim _35_I3B_UNIQUE_NO As String = CIFInfo(ClientFields.cifNO)    'FixedLengthString(CIFInfo(ClientFields.cifNO), 25, "Left", " ")
        Dim _36_FICA_STATUS As String = FICAStatus(CIFInfo(ClientFields.cifNO), SARSForm.PeriodEnd)
        Dim _37_NAME As String
        Dim _38_INITIALS As String = FixNull(CIFInfo(ClientFields.cifiInitials))
        Dim _39_FIRST_NAMES As String = FixNull(CIFInfo(ClientFields.cifiFirstName))
        'Dim _143_TRADING_NAME As String = ""
        Dim _40_ID_TYPE As String = "009"
        Dim _41_ID_NO As String = "NOIDNUMBER"
        'Dim _44_PASSPORT_NO As String = ""
        Dim _42_PASSPORT_COUNTRY As String = ""
        Dim _43_TAX_REFERENCE As String = ""
        'Dim _44_OTHER_REGISTRATION_NO As String = ""
        Dim _45_DOB As String = ""
        Dim _46_RESIDENCE_INDICATOR As String = ""
        Dim _47_NATURE_OF_PERSON As String
        'Const _48_PARTNERSHIP As String = "N"
        Dim _49_ADDR_UNIT_NO As String = ""
        Dim _50_ADDR_COMPLEX As String = ""
        Dim _51_ADDR_STREET_NO As String = ""
        Dim _52_ADDR_STREET_NAME As String = ""
        Dim _53_ADDR_SUBURB As String = ""
        Dim _54_ADDR_CITY As String = ""
        Dim _55_ADDR_POST_CODE As String = ""
        Dim _56_POSTAL_EQ_RESIDENTIAL As String = "N"
        Dim _57_PADDR_LINE_1 As String
        Dim _58_PADDR_LINE_2 As String
        Dim _59_PADDR_LINE_3 As String
        Dim _60_PADDR_LINE_4 As String
        Dim _61_PADDR_POST_CODE As String
        Dim _400_TELEPHONE_NO As String
        Dim _401_CELLPHONE_NO As String
        Dim _402_EMAIL_ADDRESS As String

        Dim PostAddress(8) As String
        Dim PhysicalAddress(8) As String
        Dim CurrentAddress(8) As String
        Dim ContactDetails(3) As String

        PostAddress = GetAddress(CIFInfo(ClientFields.cifId), CIFInfo(ClientFields.cifNO), "Postal")
        PhysicalAddress = GetAddress(CIFInfo(ClientFields.cifId), CIFInfo(ClientFields.cifNO), "Physical")
        ContactDetails = GetContactDetails(CIFInfo(ClientFields.cifId), CIFInfo(ClientFields.cifNO))

        If PhysicalAddress(AddressFields.Status) = "Address Found" Then
            CurrentAddress = PhysicalAddress
        ElseIf PostAddress(AddressFields.Status) = "Address Found" Then
            CurrentAddress = PostAddress
        Else
            CurrentAddress(AddressFields.Status) = "No Address"
            WriteExceptionToDatabase(IT3sExceptionTable, CIFInfo(ClientFields.cifNO), "Address", "Client has no addresses")
        End If

        If CurrentAddress(AddressFields.Status) = "No Address" Then
            CurrentAddress(AddressFields.Status) = "Unknown Address"
            CurrentAddress(AddressFields.Line1) = "UNKNOWN"
            CurrentAddress(AddressFields.Line2) = ""
            CurrentAddress(AddressFields.Line3) = ""
            CurrentAddress(AddressFields.Line4) = ""
            CurrentAddress(AddressFields.Suburb) = ""
            CurrentAddress(AddressFields.City) = ""
            CurrentAddress(AddressFields.PostCode) = ""
            CurrentAddress(AddressFields.Province) = ""
        End If


        Select Case CIFInfo(ClientFields.clnttype)

            Case 1   'Individuals
                _37_NAME = CIFInfo(ClientFields.cifiSurname)

                If ClientHasReferences Then
                    If FixNull(CIFInfo(ClientFields.cifiIDNO)).ToString.Length > 0 Then
                        _40_ID_TYPE = "001"
                        _41_ID_NO = IDNoModulusCheck(CIFInfo(ClientFields.cifNO), CIFInfo(ClientFields.cifiIDNO).ToString.Replace(" ", ""))
                    Else
                        If FixNull(CIFInfo(ClientFields.cifiForeignIDNO)).ToString.Length > 0 Then
                            _40_ID_TYPE = "002"
                            _41_ID_NO = CIFInfo(ClientFields.cifiForeignIDNO)
                            _42_PASSPORT_COUNTRY = SARSCountryCode(CDbl(FixNull(CIFInfo(ClientFields.ctyIdIndividual))))
                        Else
                            If FixNull(CIFInfo(ClientFields.cifiPassportNO)).ToString.Length > 0 Then
                                _40_ID_TYPE = "003"
                                _41_ID_NO = CIFInfo(ClientFields.cifiPassportNO)
                            End If
                        End If
                    End If
                End If

                If FixNull(CIFInfo(ClientFields.cifiTaxNO)).ToString.Length > 0 Then
                    _43_TAX_REFERENCE = ValidTaxReference(CIFInfo(ClientFields.cifNO), FixedLengthString(CIFInfo(ClientFields.cifiTaxNO), 10, "Left", "0", "/", " "))
                End If
                _45_DOB = FixNull(CIFInfo(ClientFields.cifiDOB))
                If _45_DOB.Length = 8 Then
                    If _41_ID_NO <> "NOIDNUMBER" Then
                        _45_DOB = ValidateDOB(CIFInfo(ClientFields.cifNO), _45_DOB, _41_ID_NO)
                    End If
                    If IsDate("#" & _45_DOB.Substring(4, 2) & "/" & _45_DOB.Substring(6, 2) & "/" & _45_DOB.Substring(0, 4) & "#") Then
                        _45_DOB = DateWithDashes(_45_DOB)
                    Else
                        _45_DOB = "0001-01-01"
                    End If
                Else
                    WriteExceptionToDatabase(IT3sExceptionTable, CIFInfo(ClientFields.cifNO), "Date of Birth", "Date of Birth " & _45_DOB & " has invalid length")
                    _45_DOB = "0001-01-01"
                End If
                If FixNull(CIFInfo(ClientFields.ctyIdIndividual)) = "" Then
                    _46_RESIDENCE_INDICATOR = "Y"
                Else
                    If CDbl(FixNull(CIFInfo(ClientFields.ctyIdIndividual))) = 510 Then
                        _46_RESIDENCE_INDICATOR = "Y"
                    Else
                        _46_RESIDENCE_INDICATOR = "N"
                    End If
                End If
                _47_NATURE_OF_PERSON = "INDIVIDUAL"
                    'If CurrentAddress(0) <> "No Address" Then
                    '    _57_PADDR_LINE_1 = CurrentAddress(AddressFields.Line1)
                    '    _58_PADDR_LINE_2 = CurrentAddress(AddressFields.Line2)
                    '    _59_PADDR_LINE_3 = CurrentAddress(AddressFields.Line3)
                    '    _60_PADDR_LINE_4 = CurrentAddress(AddressFields.City)
                    '    _61_PADDR_POST_CODE = CurrentAddress(AddressFields.PostCode)
                    'Else
                    '    _57_PADDR_LINE_1 = ""
                    '    _58_PADDR_LINE_2 = ""
                    '    _59_PADDR_LINE_3 = ""
                    '    _60_PADDR_LINE_4 = ""
                    '    _61_PADDR_POST_CODE = ""
                    'End If

            Case 2   'Companies
                _37_NAME = CIFInfo(ClientFields.cifcName)
                If FixNull(CIFInfo(ClientFields.cifcTaxNO)).ToString.Length > 0 Then
                    _43_TAX_REFERENCE = ValidTaxReference(CIFInfo(ClientFields.cifNO), FixedLengthString(CIFInfo(ClientFields.cifcTaxNO), 10, "Left", "0", "/", " "))
                End If

                '_143_TRADING_NAME = FixNull(CIFInfo(ClientFields.cifcTradingAs))

                _47_NATURE_OF_PERSON = SARSClientType(CIFInfo(ClientFields.companytype))

                If ClientHasReferences Then
                    Select Case _47_NATURE_OF_PERSON
                        Case "PRIVATE_CO", "PUBLIC_CO", "OTHER_CO", "CLOSE_CORPORATION"
                            _40_ID_TYPE = "004"
                            '_44_OTHER_REGISTRATION_NO = ValidCompanyRegNo(CIFInfo(ClientFields.cifNO),  DropChar(DropChar(FixNull(CIFInfo(ClientFields.cifcRegNO)), "/"), " "))
                            _41_ID_NO = ValidCompanyRegNo(CIFInfo(ClientFields.cifNO), DropChar(FixNull(CIFInfo(ClientFields.cifcRegNO)), " "))
                        Case "INTERVIVOS_TRUST"
                            _40_ID_TYPE = "007"
                            '_44_OTHER_REGISTRATION_NO = ValidTrustRegNo(CIFInfo(ClientFields.cifNO), DropChar(DropChar(FixNull(CIFInfo(ClientFields.cifcRegNO)), "/"), " "))
                            _41_ID_NO = DropChar(FixNull(CIFInfo(ClientFields.cifcRegNO)), " ")
                        Case Else
                            _41_ID_NO = ""
                    End Select
                Else
                    _41_ID_NO = DropChar(FixNull(CIFInfo(ClientFields.cifcRegNO)), " ")
                End If

                'If _41_ID_NO = "" Then
                '    _40_ID_TYPE = "009"
                'End If

        End Select

        _37_NAME = _37_NAME.TrimStart(" ")
        _38_INITIALS = _38_INITIALS.TrimStart(" ")
        If _38_INITIALS.Length > 30 Then
            _38_INITIALS = _38_INITIALS.Substring(0, 30)
        End If

        If CIFInfo(ClientFields.clnttype) <> 1 Then
            Select Case CIFInfo(ClientFields.companytype)
                Case 15, 16, 25, 28
                    _39_FIRST_NAMES = _37_NAME
            End Select
        End If

        If _41_ID_NO = "" Then
            _40_ID_TYPE = "009"
            _41_ID_NO = "NOIDNUMBER"
        End If

        _33_UNIQUE_NO = GetExistingUniqueNo(_35_I3B_UNIQUE_NO, _31_RECORD_TYPE, SARSForm.TestFile)
        If _33_UNIQUE_NO <> "Unique_No" Then
            If RecordCompletelyRejected(_33_UNIQUE_NO, SARSForm.TestFile) Then
                _32_RECORD_STATUS = "N"
            Else
                '_34_ROW_NO = _33_UNIQUE_NO
                _32_RECORD_STATUS = "C"
            End If
        Else
            _32_RECORD_STATUS = "N"
        End If

        If CurrentAddress(0) <> "No Address" Then
            _57_PADDR_LINE_1 = CurrentAddress(AddressFields.Line1)
            _58_PADDR_LINE_2 = CurrentAddress(AddressFields.Line2)
            _59_PADDR_LINE_3 = CurrentAddress(AddressFields.Line3)
            _60_PADDR_LINE_4 = CurrentAddress(AddressFields.City)
            _61_PADDR_POST_CODE = CurrentAddress(AddressFields.PostCode)
        Else
            _57_PADDR_LINE_1 = ""
            _58_PADDR_LINE_2 = ""
            _59_PADDR_LINE_3 = ""
            _60_PADDR_LINE_4 = ""
            _61_PADDR_POST_CODE = ""
        End If

        If _61_PADDR_POST_CODE = "" Then
            _61_PADDR_POST_CODE = "0000"
        End If

        _400_TELEPHONE_NO = ContactDetails(0)
        If ContactDetails(1) = "" Then
            _401_CELLPHONE_NO = "999999999999999"
        Else
            _401_CELLPHONE_NO = ContactDetails(1)
        End If
        If ContactDetails(2) = "" Then
            _402_EMAIL_ADDRESS = "NO@EMAIL.COM"
        Else
            _402_EMAIL_ADDRESS = ContactDetails(2)
        End If

        PersonalDetailsRecord = _30_SEC_ID & DataDelimiter &
                                _31_RECORD_TYPE & DataDelimiter &
                                _32_RECORD_STATUS & DataDelimiter &
                                _33_UNIQUE_NO & DataDelimiter &
                                _34_ROW_NO & DataDelimiter &
                                _35_I3B_UNIQUE_NO & DataDelimiter &
                                _36_FICA_STATUS & DataDelimiter &
                                _37_NAME & DataDelimiter &
                                _38_INITIALS & DataDelimiter &
                                _39_FIRST_NAMES & DataDelimiter &
                                _40_ID_TYPE & DataDelimiter &
                                _41_ID_NO & DataDelimiter &
                                _42_PASSPORT_COUNTRY & DataDelimiter &
                                _43_TAX_REFERENCE & DataDelimiter &
                                _45_DOB & DataDelimiter &
                                _46_RESIDENCE_INDICATOR & DataDelimiter &
                                _47_NATURE_OF_PERSON & DataDelimiter &
                                _49_ADDR_UNIT_NO & DataDelimiter &
                                _50_ADDR_COMPLEX & DataDelimiter &
                                _51_ADDR_STREET_NO & DataDelimiter &
                                _52_ADDR_STREET_NAME & DataDelimiter &
                                _53_ADDR_SUBURB & DataDelimiter &
                                _54_ADDR_CITY & DataDelimiter &
                                _55_ADDR_POST_CODE & DataDelimiter &
                                _56_POSTAL_EQ_RESIDENTIAL & DataDelimiter &
                                _57_PADDR_LINE_1 & DataDelimiter &
                                _58_PADDR_LINE_2 & DataDelimiter &
                                _59_PADDR_LINE_3 & DataDelimiter &
                                _60_PADDR_LINE_4 & DataDelimiter &
                                _61_PADDR_POST_CODE & DataDelimiter &
                                _400_TELEPHONE_NO & DataDelimiter &
                                _401_CELLPHONE_NO & DataDelimiter &
                                _402_EMAIL_ADDRESS

        'PersonalDetailsRecord = _30_SEC_ID & DataDelimiter & _
        '            _31_RECORD_TYPE & DataDelimiter & _
        '            _32_RECORD_STATUS & DataDelimiter & _
        '            _33_UNIQUE_NO & DataDelimiter & _
        '            _34_ROW_NO & DataDelimiter & _
        '            _35_I3B_UNIQUE_NO & DataDelimiter & _
        '            _36_FICA_STATUS & DataDelimiter & _
        '            _37_NAME & DataDelimiter & _
        '            _38_INITIALS & DataDelimiter & _
        '            _39_FIRST_NAMES & DataDelimiter & _
        '            _143_TRADING_NAME & DataDelimiter & _
        '            _40_ID_TYPE & DataDelimiter & _
        '            _41_ID_NO & DataDelimiter & _
        '            _42_PASSPORT_COUNTRY & DataDelimiter & _
        '            _43_TAX_REFERENCE & DataDelimiter & _
        '            _45_DOB & DataDelimiter & _
        '            _46_RESIDENCE_INDICATOR & DataDelimiter & _
        '            _47_NATURE_OF_PERSON & DataDelimiter & _
        '            _48_PARTNERSHIP & DataDelimiter & _
        '            _49_ADDR_UNIT_NO & DataDelimiter & _
        '            _50_ADDR_COMPLEX & DataDelimiter & _
        '            _51_ADDR_STREET_NO & DataDelimiter & _
        '            _52_ADDR_STREET_NAME & DataDelimiter & _
        '            _53_ADDR_SUBURB & DataDelimiter & _
        '            _54_ADDR_CITY & DataDelimiter & _
        '            _55_ADDR_POST_CODE & DataDelimiter & _
        '            _56_POSTAL_EQ_RESIDENTIAL & DataDelimiter & _
        '            _57_PADDR_LINE_1 & DataDelimiter & _
        '            _58_PADDR_LINE_2 & DataDelimiter & _
        '            _59_PADDR_LINE_3 & DataDelimiter & _
        '            _60_PADDR_LINE_4 & DataDelimiter & _
        '            _61_PADDR_POST_CODE

        'End If
    End Function

    'Private Function PersonalDetailsRecord(ByVal CIFInfo As Object) As String

    '    Const SEC_ID As String = "P"

    '    Dim IT3_PERS_ID As String
    '    Dim IT_REF_NO As String
    '    Dim PERIOD_START As String = SARSForm.PeriodStart(SARSForm.txtPeriodLength.Text)
    '    Dim PERIOD_END As String = SARSForm.PeriodEnd
    '    Dim TP_CATEGORY As String
    '    Dim TP_ID As String
    '    Dim TP_OTHER_ID As String
    '    Dim CO_REG_NO As String
    '    Dim TRUST_DEED_NO As String
    '    Dim TP_NAME As String
    '    Dim TP_INITS As String
    '    Dim TP_FIRSTNAMES As String
    '    Dim TP_DOB As String
    '    Dim TP_TRADE_NAME As String
    '    Dim TP_POST_ADDR(5) As String
    '    'Dim TP_POST_CODE As String
    '    Dim TP_PHY_ADDR(5) As String
    '    'Dim TP_PHY_CODE As String
    '    Dim TP_SA_RES As String
    '    Dim PARTNERSHIP As String

    '    IT3_PERS_ID = FixedLengthString(CIFInfo(ClientFields.cifNO), 25, "Left", " ")
    '    TP_POST_ADDR = GetAddress(CIFInfo(ClientFields.cifId), "Postal")
    '    TP_PHY_ADDR = GetAddress(CIFInfo(ClientFields.cifId), "Physical")
    '    PARTNERSHIP = "N"

    '    Select Case CIFInfo(ClientFields.clnttype)

    '        Case 1   'Individuals
    '            IT_REF_NO = ValidTaxReference(CIFInfo(ClientFields.cifNO), FixedLengthString(CIFInfo(ClientFields.cifiTaxNO), 10, "Left", " ", "/"))
    '            TP_CATEGORY = "01"
    '            TP_ID = FixedLengthString(CIFInfo(ClientFields.cifiIDNO), 13, "Left", " ")
    '            If FixNull(CIFInfo(ClientFields.cifiForeignIDNO)) <> "" Then
    '                TP_OTHER_ID = FixedLengthString(CIFInfo(ClientFields.cifiForeignIDNO), 10, "Left", " ")
    '            Else
    '                If FixNull(CIFInfo(ClientFields.cifiPassportNO)) <> "" Then
    '                    TP_OTHER_ID = FixedLengthString(CIFInfo(ClientFields.cifiPassportNO), 10, "Left", " ")
    '                Else
    '                    TP_OTHER_ID = FixedLengthString(" ", 10, "Left", " ")
    '                End If
    '            End If
    '            CO_REG_NO = FixedLengthString(" ", 15, "Left", " ")
    '            TRUST_DEED_NO = FixedLengthString(" ", 10, "Left", " ")
    '            TP_NAME = FixedLengthString(CIFInfo(ClientFields.cifiSurname), 120, "Left", " ")
    '            TP_INITS = FixedLengthString(CIFInfo(ClientFields.cifiInitials), 5, "Left", " ")
    '            TP_FIRSTNAMES = FixedLengthString(CIFInfo(ClientFields.cifiFirstName), 90, "Left", " ")
    '            TP_DOB = FixedLengthString(CIFInfo(ClientFields.cifiDOB), 8, "Left", " ")
    '            TP_TRADE_NAME = FixedLengthString(" ", 120, "Left", " ")
    '            If FixNull(CIFInfo(ClientFields.ctyIdIndividual)) = "" Then
    '                TP_SA_RES = "Y"
    '            Else
    '                If CDbl(FixNull(CIFInfo(ClientFields.ctyIdIndividual))) = 510 Then
    '                    TP_SA_RES = "Y"
    '                Else
    '                    TP_SA_RES = "N"
    '                End If
    '            End If
    '        Case 2   'Companies
    '            IT_REF_NO = ValidTaxReference(CIFInfo(ClientFields.cifNO), FixedLengthString(CIFInfo(ClientFields.cifcTaxNO), 10, "Left", " ", "/"))
    '            Select Case CIFInfo(ClientFields.companytype)
    '                Case 18, 19, 26
    '                    TP_CATEGORY = "03"
    '                Case Else
    '                    TP_CATEGORY = "02"
    '            End Select
    '            TP_ID = FixedLengthString(" ", 13, "Left", " ")
    '            TP_OTHER_ID = FixedLengthString(" ", 10, "Left", " ")
    '            If TP_CATEGORY = "03" Then
    '                CO_REG_NO = FixedLengthString(" ", 15, "Left", " ")
    '                TRUST_DEED_NO = FixedLengthString(CIFInfo(ClientFields.cifcRegNO), 10, "Left", " ")
    '            Else
    '                CO_REG_NO = FixedLengthString(CIFInfo(ClientFields.cifcRegNO), 15, "Left", " ")
    '                TRUST_DEED_NO = FixedLengthString(" ", 10, "Left", " ")
    '            End If
    '            TP_NAME = FixedLengthString(CIFInfo(ClientFields.cifcName), 120, "Left", " ")
    '            TP_INITS = FixedLengthString(" ", 5, "Left", " ")
    '            TP_FIRSTNAMES = FixedLengthString(" ", 90, "Left", " ")
    '            TP_DOB = FixedLengthString(" ", 8, "Left", " ")
    '            TP_TRADE_NAME = FixedLengthString(CIFInfo(ClientFields.cifcTradingAs), 120, "Left", " ")
    '            If CIFInfo(ClientFields.ctyIdCompany) = 510 Then
    '                TP_SA_RES = "Y"
    '            Else
    '                TP_SA_RES = "N"
    '            End If
    '    End Select

    '    PersonalDetailsRecord = SEC_ID & _
    '                            IT3_PERS_ID & _
    '                            IT_REF_NO & _
    '                            PERIOD_START & _
    '                            PERIOD_END & _
    '                            TP_CATEGORY & _
    '                            TP_ID & _
    '                            TP_OTHER_ID & _
    '                            CO_REG_NO & _
    '                            TRUST_DEED_NO & _
    '                            TP_NAME & _
    '                            TP_INITS & _
    '                            TP_FIRSTNAMES & _
    '                            TP_DOB & _
    '                            TP_TRADE_NAME & _
    '                            TP_POST_ADDR(1) & _
    '                            TP_POST_ADDR(2) & _
    '                            TP_POST_ADDR(3) & _
    '                            TP_POST_ADDR(4) & _
    '                            TP_POST_ADDR(5) & _
    '                            TP_PHY_ADDR(1) & _
    '                            TP_PHY_ADDR(2) & _
    '                            TP_PHY_ADDR(3) & _
    '                            TP_PHY_ADDR(4) & _
    '                            TP_PHY_ADDR(5) & _
    '                            TP_SA_RES & _
    '                            PARTNERSHIP
    'End Function

    Private Function FICAStatus(ByVal cifNo As String, ByVal strPeriodEnd As String) As String

        Dim strSQL As String = "SELECT   cifFICAID " &
                               "FROM     tblCIF " &
                               "WHERE    cifNO = '" & cifNo & "' "

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)
        Dim FICA As String = "N"

        cn.Open()
        Try
            drSQL = cmd.ExecuteReader()
            drSQL.Read()
            If drSQL.Item(0) Then
                If drSQL(0) < 3 Then
                    FICA = "Y"
                End If
            End If
        Catch ex As Exception
        End Try
        cn.Close()
        If FICA = "N" Then
            WriteExceptionToDatabase(IT3sExceptionTable, cifNo, "FICA", "Client is not FICA compliant")
        End If
        Return FICA
    End Function

    'Private Function FICAStatus(ByVal cifNo As String, ByVal strPeriodEnd As String) As String

    '    Dim strSQL As String = "SELECT   cifFICA " &
    '                           "FROM     tblCIF " &
    '                           "WHERE    cifNO = '" & cifNo & "' "

    '    Dim cn As SqlConnection = New SqlConnection(strConnection)
    '    Dim drSQL As SqlDataReader
    '    Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)
    '    Dim FICA As String = "N"

    '    cn.Open()
    '    Try
    '        drSQL = cmd.ExecuteReader()
    '        drSQL.Read()
    '        If drSQL.Item(0) Then
    '            FICA = "Y"
    '        End If
    '    Catch ex As Exception
    '    End Try
    '    cn.Close()
    '    If FICA = "N" Then
    '        WriteExceptionToDatabase(IT3sExceptionTable, cifNo, "FICA", "Client is not FICA compliant")
    '    End If
    '    Return FICA
    'End Function

    'Private Function FICAStatus(ByVal cifNo As String, ByVal strPeriodEnd As String) As String
    '    'This is the old method, using a document count.  For the purposes of FICA2016, we can simply use the CIF flag
    '    Dim strSQL As String = "SELECT   CIFs_With_Open_Accounts.NoOfAccounts - CIFs_With_Open_Accounts.SchemeAccounts - CIFs_With_Open_Accounts.CorporateAccounts AS NoOfAccounts, " & _
    '                                    "SUM(FICA_Docs.IndActualReq + FICA_Docs.CoActualReq) - SUM(FICA_Docs.IndDocs + FICA_Docs.CoDocs) AS DocsOutstanding " & _
    '                           "FROM     tblCIF AS tblCIF LEFT OUTER JOIN " & _
    '                                    "tblCIFCompany AS tblCIFCompany ON tblCIF.cifId = tblCIFCompany.cifId LEFT OUTER JOIN " & _
    '                                    "tblCIFIndividual AS tblCIFIndividual ON tblCIF.cifId = tblCIFIndividual.cifId LEFT OUTER JOIN " & _
    '                                    "FICA_Docs AS FICA_Docs ON tblCIF.cifId = FICA_Docs.cifId LEFT OUTER JOIN " & _
    '                                    "CIFs_With_Open_Accounts AS CIFs_With_Open_Accounts ON tblCIF.cifId = CIFs_With_Open_Accounts.cifId LEFT OUTER JOIN " & _
    '                                    "FICA_Docs_Total_Required AS FICA_Docs_Total_Required ON tblCIF.cifId = FICA_Docs_Total_Required.cifId LEFT OUTER JOIN " & _
    '                                    "mblFICADocGroup AS mblFICADocGroup ON FICA_Docs.GroupId = mblFICADocGroup.ficDocGrpId LEFT OUTER JOIN " & _
    '                                    "mblCompanyType AS mblCompanyType ON tblCIFCompany.ctId = mblCompanyType.ctId " & _
    '                           "WHERE    tblCIF.cifNO = '" & cifNo & "' " & _
    '                           "GROUP BY CIFs_With_Open_Accounts.NoOfAccounts - CIFs_With_Open_Accounts.SchemeAccounts - CIFs_With_Open_Accounts.CorporateAccounts"

    '    Dim cn As SqlConnection = New SqlConnection(strConnection)
    '    Dim drSQL As SqlDataReader
    '    Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)
    '    Dim Done As Boolean = False
    '    Dim Tries As Byte = 0
    '    Dim FICA As String = "N"

    '    cn.Open()
    '    While Not Done And Tries < 3
    '        Try
    '            drSQL = cmd.ExecuteReader()
    '            drSQL.Read()
    '            If drSQL.Item(0) <= 0 Or drSQL.Item(1) <= 0 Then
    '                FICA = "Y"
    '            End If
    '            Done = True
    '        Catch ex As Exception
    '            Tries += 1
    '        End Try
    '    End While
    '    cn.Close()
    '    Select Case Done
    '        Case True
    '            If FICA = "N" Then
    '                WriteExceptionToDatabase(IT3sExceptionTable, cifNo, "FICA", "Client is not FICA compliant")
    '            End If
    '        Case False
    '            FICA = "N"
    '            WriteExceptionToDatabase(IT3sExceptionTable, cifNo, "FICA", "Could not determine FICA status of client - assuming not compliant")
    '    End Select
    '    Return FICA
    'End Function

    'Private Function FICAStatus(ByVal cifNo As String, ByVal strPeriodEnd As String) As String

    '    Dim strSQL As String = "SELECT   TOP (1) tblAccountDailyBalBak.FICADocsOutstanding, tblAccountDailyBalBak.PostDate " & _
    '                           "FROM     tblAccountDailyBalBak INNER JOIN " & _
    '                                    "jblCIFAccount ON tblAccountDailyBalBak.accId = jblCIFAccount.accId INNER JOIN " & _
    '                                    "tblCIF ON jblCIFAccount.cifId = tblCIF.cifId " & _
    '                           "WHERE    tblCIF.cifNO = '" & cifNo & "' AND " & _
    '                                    "tblAccountDailyBalBak.batchMonth >= '" & strPeriodEnd.Substring(0, 6) & "01' " & _
    '                           "ORDER BY tblAccountDailyBalBak.PostDate DESC"

    '    Dim strSQL1 As String = "SELECT cifFICA " & _
    '                            "FROM   tblCIF " & _
    '                            "WHERE  cifNO = '" & cifNo & "'"

    '    Dim cn As SqlConnection = New SqlConnection(strConnection)
    '    Dim drSQL As SqlDataReader
    '    Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)
    '    Dim cmd1 As SqlCommand = New SqlCommand(strSQL1, cn)
    '    Dim Done As Boolean = False
    '    Dim Tries As Byte = 0
    '    Dim FICA As String = "N"

    '    cn.Open()
    '    While Not Done And Tries < 3
    '        Try
    '            drSQL = cmd.ExecuteReader()
    '            drSQL.Read()
    '            If drSQL.Item(0) <= 0 Then
    '                FICA = "Y"
    '            End If
    '            Done = True
    '        Catch ex As Exception
    '            Tries += 1
    '        End Try
    '    End While
    '    cn.Close()
    '    Select Case Done
    '        Case True
    '            If FICA = "N" Then
    '                WriteExceptionToDatabase(IT3sExceptionTable, cifNo, "FICA", "Client is not FICA compliant")
    '            End If
    '        Case False
    '            cn.Open()
    '            Try
    '                drSQL = cmd1.ExecuteReader()
    '                drSQL.Read()
    '                If drSQL.Item(0) Then
    '                    FICA = "Y"
    '                Else
    '                    FICA = "N"
    '                    WriteExceptionToDatabase(IT3sExceptionTable, cifNo, "FICA", "Client is not FICA compliant")
    '                End If
    '            Catch ex As Exception
    '                FICA = "N"
    '                WriteExceptionToDatabase(IT3sExceptionTable, cifNo, "FICA", "Could not determine FICA status of client - assuming not compliant")
    '            End Try
    '            cn.Close()
    '    End Select
    '    Return FICA
    'End Function

    Private Function SARSCountryCode(ByVal ctyCode As Integer) As String

        Dim strSQL As String = "SELECT ctyCode " &
                               "FROM   mblCountry " &
                               "WHERE  ctyId = " & ctyCode

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)

        cn.Open()
        drSQL = cmd.ExecuteReader()
        Try
            drSQL.Read()
            SARSCountryCode = drSQL.Item(0)
        Catch ex As Exception
            SARSCountryCode = NoCountry
        End Try
        cn.Close()
    End Function

    Private Function SARSAccountCode(ByVal atId As Integer) As String

        Dim strSQL As String = "SELECT atSARSCode " &
                               "FROM   mblAccountType " &
                               "WHERE  atId = " & atId

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)

        cn.Open()
        drSQL = cmd.ExecuteReader()
        Try
            drSQL.Read()
            SARSAccountCode = drSQL.Item(0)
        Catch ex As Exception
            SARSAccountCode = SARSOtherAccountCode
        End Try
        cn.Close()
    End Function

    Private Function SARSClientType(ByVal CompanyType As Integer) As String

        Dim strSQL As String = "SELECT ctSARSCode " &
                               "FROM   mblCompanyType " &
                               "WHERE  ctId = " & CompanyType

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)

        cn.Open()
        drSQL = cmd.ExecuteReader()
        Try
            drSQL.Read()
            SARSClientType = drSQL.Item(0)
        Catch ex As Exception
            SARSClientType = "INDIVIDUAL"
        End Try
        cn.Close()
        Return SARSClientType
    End Function

    Private Function FirstPosting(ByVal accNo As String, ByVal AsAtDate As String) As Double

        Dim Success As Boolean = False
        Dim Tries As Byte = 0

        Dim strSQL As String = "SELECT TOP (1) trnAmt " &
                               "FROM   tblAccount INNER JOIN tblTransactions " &
                                                   "ON tblAccount.accId = tblTransactions.accId " &
                               "WHERE  accNO = '" & accNo & "' AND " &
                                      "trnTransactionDate >= '" & AsAtDate & "' AND " &
                                      "trnTransactionDate <= '" & SARSForm.PeriodEnd & "' AND " &
                                      "trnAmt <> 0 AND " &
                                      "tblTransactions.Deleted = 0 " &
                               "ORDER BY trnId ASC"

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)

        While Not Success And Tries < 3
            Try
                Tries += 1
                cn.Open()
                drSQL = cmd.ExecuteReader()
                Success = True
            Catch ex As Exception
            End Try
        End While
        Try
            drSQL.Read()
            FirstPosting = drSQL.Item(0)
        Catch ex As Exception
            FirstPosting = 0
        End Try
        cn.Close()
    End Function

    Private Function AccountOpeningBalance(ByVal accNo As String, ByVal AsAtDate As String) As Double

        Dim tmpAccountOpeningBalance As Double = 0
        Dim Success As Boolean = False
        Dim Tries As Byte = 0

        Dim strSQL As String = "SELECT TOP (1) trnTransactionDate, trnAmt, trnAccountBal " &
                               "FROM   tblAccount INNER JOIN tblTransactions " &
                                                 "ON tblAccount.accId = tblTransactions.accId " &
                               "WHERE  accNO = '" & accNo & "' AND " &
                                      "trnTransactionDate <= '" & AsAtDate & "' AND " &
                                      "tblTransactions.Deleted = 0 " &
                               "ORDER BY trnId DESC"

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand

        cmd = New SqlCommand(strSQL, cn)
        While Not Success And Tries < 3
            Try
                Tries += 1
                cn.Open()
                drSQL = cmd.ExecuteReader()
                Success = True
            Catch ex As Exception
            End Try
        End While
        If Success Then
            Try
                drSQL.Read()
                tmpAccountOpeningBalance = drSQL.Item(2)
                If drSQL.Item(0) = AsAtDate And Not (drSQL.Item(1) = tmpAccountOpeningBalance) Then
                    tmpAccountOpeningBalance = tmpAccountOpeningBalance - drSQL.Item(1)
                End If
            Catch ex As Exception
                tmpAccountOpeningBalance = 0
            End Try
            cn.Close()
        End If
        'If tmpAccountOpeningBalance = 0 Then
        '    tmpAccountOpeningBalance = FirstPosting(accNo, AsAtDate)
        'End If
        tmpAccountOpeningBalance = FixNullNum(tmpAccountOpeningBalance)
        TotalMoney = TotalMoney + tmpAccountOpeningBalance
        Return tmpAccountOpeningBalance
    End Function

    Private Function AccountClosingBalance(ByVal accNO As String, ByVal AsAtDate As String) As Double

        Dim tmpAccountClosingBalance As Double = 0
        Dim Success As Boolean = False
        Dim Tries As Byte = 0

        Dim strSQL As String = "SELECT TOP (1) trnTransactionDate, trnAmt, trnAccountBal " &
                           "FROM   tblAccount INNER JOIN tblTransactions " &
                                                  "ON tblAccount.accId = tblTransactions.accId " &
                           "WHERE  accNO = '" & accNO & "' AND " &
                                  "trnTransactionDate <= '" & AsAtDate & "' AND " &
                                  "tblTransactions.Deleted = 0 " &
                           "ORDER BY trnId DESC"

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand

        cmd = New SqlCommand(strSQL, cn)
        While Not Success And Tries < 3
            Try
                Tries += 1
                cn.Open()
                drSQL = cmd.ExecuteReader()
                Success = True
            Catch ex As Exception
            End Try
        End While
        Try
            drSQL.Read()
            tmpAccountClosingBalance = FixNullNum(drSQL.Item(2))
        Catch ex As Exception
        End Try
        cn.Close()
        TotalMoney = TotalMoney + System.Math.Abs(tmpAccountClosingBalance)
        Return tmpAccountClosingBalance
    End Function

    Private Function AccountCloseDate(ByVal accNO As String) As String

        Dim Success As Boolean = False
        Dim Tries As Byte = 0

        Dim strSQL As String = "SELECT TOP (1) trnTransactionDate " &
                               "FROM   tblAccount INNER JOIN tblTransactions " &
                                                  "ON tblAccount.accId = tblTransactions.accId " &
                               "WHERE  accNO = '" & accNO & "' AND " &
                                      "tblTransactions.Deleted = 0 " &
                               "ORDER BY trnId DESC"

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand

        cmd = New SqlCommand(strSQL, cn)
        While Not Success And Tries < 3
            Try
                Tries += 1
                cn.Open()
                drSQL = cmd.ExecuteReader()
                Success = True
            Catch ex As Exception
            End Try
        End While
        Try
            drSQL.Read()
            AccountCloseDate = FixNull(drSQL.Item(0))
        Catch ex As Exception
            AccountCloseDate = ""
        End Try
        cn.Close()
    End Function

    'Private Function TotalMonetaryValue() As String

    'Dim strSQL As String = "SELECT TOP (1) trnTransactionDate " & _
    '                       "FROM   tblAccount INNER JOIN tblTransactions " & _
    '                                          "ON tblAccount.accId = tblTransactions.accId " & _
    '                       "WHERE  accNO = '" & accNO & "' AND " & _
    '                              "tblTransactions.Deleted = 0 " & _
    '                       "ORDER BY trnId DESC"

    'Dim cn As SqlConnection = New SqlConnection(strConnection)
    'Dim drSQL As SqlDataReader
    'Dim cmd As SqlCommand

    'cmd = New SqlCommand(strSQL, cn)
    'cn.Open()
    'drSQL = cmd.ExecuteReader()
    'Try
    '    drSQL.Read()
    '    AccountCloseDate = FixNull(drSQL.Item(0))
    'Catch ex As Exception
    '    AccountCloseDate = ""
    'End Try
    'cn.Close()
    'End Function

    '    Private Function IncomeDetails(ByVal AccountInfo As Object, ByVal InterestIncome As Double) As String
    Private Function IncomeDetails(ByVal AccountInfo As Object) As String

        'Dim AllTransactions(12, 2) As String
        'AllTransactions = AccountTransactions(AccountInfo(AccountFields.taxAccountNo))

        Const _91_SEC_ID As String = "B"
        Const _92_RECORD_TYPE As String = "AHFD"
        Dim _93_RECORD_STATUS As String
        Dim _94_UNIQUE_NO As String = "Unique_No"  'TotalRecords.ToString
        Dim _95_ROW_NO As String = "Rec_No"
        Dim _96_I3B_UNIQUE_NO As String = AccountInfo(AccountFields.taxCIF)    'FixedLengthString(AccountInfo(AccountFields.taxCIF), 25, "Left", " ")
        'Dim _97_INCOME_SOURCE_CODE As String = SourceCode
        Dim _98_ACCOUNT_NO As String = AccountInfo(AccountFields.taxAccountNo)
        Dim _99_ACCOUNT_TYPE As String = "19"
        Dim _901_SHARIA_INDICATOR As String = "N"
        'Dim _99_ACCOUNT_TYPE As String = SARSAccountCode(AccountInfo(AccountFields.atId))
        'Dim _100_MARCH_CREDITS As String = FormatNumber(AllTransactions(3, 2), 2, TriState.True, TriState.False, TriState.False)
        'Dim _101_APRIL_CREDITS As String = FormatNumber(AllTransactions(4, 2), 2, TriState.True, TriState.False, TriState.False)
        'Dim _102_MAY_CREDITS As String = FormatNumber(AllTransactions(5, 2), 2, TriState.True, TriState.False, TriState.False)
        'Dim _103_JUNE_CREDITS As String = FormatNumber(AllTransactions(6, 2), 2, TriState.True, TriState.False, TriState.False)
        'Dim _104_JULY_CREDITS As String = FormatNumber(AllTransactions(7, 2), 2, TriState.True, TriState.False, TriState.False)
        'Dim _105_AUGUST_CREDITS As String = FormatNumber(AllTransactions(8, 2), 2, TriState.True, TriState.False, TriState.False)
        'Dim _106_SEPTEMBER_CREDITS As String = FormatNumber(AllTransactions(9, 2), 2, TriState.True, TriState.False, TriState.False)
        'Dim _107_OCTOBER_CREDITS As String = FormatNumber(AllTransactions(10, 2), 2, TriState.True, TriState.False, TriState.False)
        'Dim _108_NOVEMBER_CREDITS As String = FormatNumber(AllTransactions(11, 2), 2, TriState.True, TriState.False, TriState.False)
        'Dim _109_DECEMBER_CREDITS As String = FormatNumber(AllTransactions(12, 2), 2, TriState.True, TriState.False, TriState.False)
        'Dim _110_JANUARY_CREDITS As String = FormatNumber(AllTransactions(1, 2), 2, TriState.True, TriState.False, TriState.False)
        'Dim _111_FEBRUARY_CREDITS As String = FormatNumber(AllTransactions(2, 2), 2, TriState.True, TriState.False, TriState.False)
        'Dim _112_MARCH_DEBITS As String = FormatNumber(AllTransactions(3, 1) * -1, 2, TriState.True, TriState.False, TriState.False)
        'Dim _113_APRIL_DEBITS As String = FormatNumber(AllTransactions(4, 1) * -1, 2, TriState.True, TriState.False, TriState.False)
        'Dim _114_MAY_DEBITS As String = FormatNumber(AllTransactions(5, 1) * -1, 2, TriState.True, TriState.False, TriState.False)
        'Dim _115_JUNE_DEBITS As String = FormatNumber(AllTransactions(6, 1) * -1, 2, TriState.True, TriState.False, TriState.False)
        'Dim _116_JULY_DEBITS As String = FormatNumber(AllTransactions(7, 1) * -1, 2, TriState.True, TriState.False, TriState.False)
        'Dim _117_AUGUST_DEBITS As String = FormatNumber(AllTransactions(8, 1) * -1, 2, TriState.True, TriState.False, TriState.False)
        'Dim _118_SEPTEMBER_DEBITS As String = FormatNumber(AllTransactions(9, 1) * -1, 2, TriState.True, TriState.False, TriState.False)
        'Dim _119_OCTOBER_DEBITS As String = FormatNumber(AllTransactions(10, 1) * -1, 2, TriState.True, TriState.False, TriState.False)
        'Dim _120_NOVEMBER_DEBITS As String = FormatNumber(AllTransactions(11, 1) * -1, 2, TriState.True, TriState.False, TriState.False)
        'Dim _121_DECEMBER_DEBITS As String = FormatNumber(AllTransactions(12, 1) * -1, 2, TriState.True, TriState.False, TriState.False)
        'Dim _122_JANUARY_DEBITS As String = FormatNumber(AllTransactions(1, 1) * -1, 2, TriState.True, TriState.False, TriState.False)
        'Dim _123_FEBRUARY_DEBITS As String = FormatNumber(AllTransactions(2, 1) * -1, 2, TriState.True, TriState.False, TriState.False)
        'Const _124_TOTAL_INTEREST_PAID As String = "0.00"
        'Dim _125_TOTAL_INTEREST_EARNED As String      '= FormatNumber(AccountInfo(AccountFields.taxPaid) + AccountInfo(AccountFields.taxAccrued), 2, TriState.True, TriState.False, TriState.False)
        Dim _126_OPENING_BALANCE As String = FormatNumber(AccountOpeningBalance(AccountInfo(AccountFields.taxAccountNo), SARSForm.PeriodStart(SARSForm.txtPeriodLength.Text)), 2, TriState.True, TriState.False, TriState.False)
        Dim _127_ACCOUNT_START_DATE As String
        Dim _128_CLOSING_BALANCE As String = FormatNumber(AccountClosingBalance(AccountInfo(AccountFields.taxAccountNo), SARSForm.PeriodEnd), 2, TriState.True, TriState.False, TriState.False)
        Dim _129_ACCOUNT_CLOSING_DATE As String
        'Dim _130_FOREIGN_TAX_PAID As String = "0.00"
        Dim _80_NET_RETURN_ON_INVESTMENT_SOURCE_CODE As String = "4239"
        Dim _81_NET_RETURN_ON_INVESTMENT As String = "0.00"
        Dim _82_INTEREST_SOURCE_CODE As String = "4241"
        Dim _83_INTEREST As String
        Dim _84_DIVIDENDS_SOURCE_CODE As String = "4242"
        Dim _85_DIVIDENDS As String = "0.00"
        Dim _86_CAPITAL_SOURCE_CODE As String = "4244"
        Dim _87_CAPITAL_GAIN As String = "0.00"
        Dim _801_OTHER_SOURCE_CODE As String = ""
        Dim _802_OTHER As String = ""
        Dim _88_MARKET_VALUE As String = _128_CLOSING_BALANCE
        Dim _89_TRANSACTION_VALUE_INDICATOR As String = "N"


        'If AccountInfo(AccountFields.taxPaid) + AccountInfo(AccountFields.taxAccrued) < 0 Then
        '    _125_TOTAL_INTEREST_EARNED = "0.00"
        'Else
        '    _125_TOTAL_INTEREST_EARNED = FormatNumber(AccountInfo(AccountFields.taxPaid) + AccountInfo(AccountFields.taxAccrued), 2, TriState.True, TriState.False, TriState.False)
        '    TotalMoney = TotalMoney + System.Math.Abs(AccountInfo(AccountFields.taxPaid) + AccountInfo(AccountFields.taxAccrued))
        'End If

        'TotalMoney = TotalMoney + _126_OPENING_BALANCE
        'TotalMoney = TotalMoney + _128_CLOSING_BALANCE

        If AccountInfo(AccountFields.accRegDate) > SARSForm.PeriodStart(SARSForm.txtPeriodLength.Text) Then
            _127_ACCOUNT_START_DATE = DateWithDashes(AccountInfo(AccountFields.accRegDate))
        Else
            _127_ACCOUNT_START_DATE = DateWithDashes(SARSForm.PeriodStart(SARSForm.txtPeriodLength.Text))
        End If

        If AccountInfo(AccountFields.asId) > 0 Then
            _129_ACCOUNT_CLOSING_DATE = AccountCloseDate(AccountInfo(AccountFields.taxAccountNo))
            If _129_ACCOUNT_CLOSING_DATE > SARSForm.PeriodEnd Then
                _129_ACCOUNT_CLOSING_DATE = DateWithDashes(SARSForm.PeriodEnd)
            Else
                _129_ACCOUNT_CLOSING_DATE = DateWithDashes(_129_ACCOUNT_CLOSING_DATE)
            End If
        Else
            _129_ACCOUNT_CLOSING_DATE = DateWithDashes(SARSForm.PeriodEnd)
        End If

        If AccountInfo(AccountFields.taxPaid) + AccountInfo(AccountFields.taxAccrued) < 0 Then
            _83_INTEREST = "0.00"
        Else
            _83_INTEREST = FormatNumber(AccountInfo(AccountFields.taxPaid) + AccountInfo(AccountFields.taxAccrued), 2, TriState.True, TriState.False, TriState.False)
            TotalMoney = TotalMoney + System.Math.Abs(AccountInfo(AccountFields.taxPaid) + AccountInfo(AccountFields.taxAccrued))
        End If

        TotalMoney = TotalMoney + _88_MARKET_VALUE

        _94_UNIQUE_NO = GetExistingUniqueNo(_96_I3B_UNIQUE_NO, _92_RECORD_TYPE, SARSForm.TestFile, _98_ACCOUNT_NO)
        If _94_UNIQUE_NO <> "Unique_No" Then
            If RecordCompletelyRejected(_94_UNIQUE_NO, SARSForm.TestFile) Then
                _93_RECORD_STATUS = "N"
            Else
                '_95_ROW_NO = _94_UNIQUE_NO
                _93_RECORD_STATUS = "C"
            End If
        Else
            _93_RECORD_STATUS = "N"
        End If

        WriteTransactionDetails(AccountInfo, _94_UNIQUE_NO)

        IncomeDetails = _91_SEC_ID & DataDelimiter &
                        _92_RECORD_TYPE & DataDelimiter &
                        _93_RECORD_STATUS & DataDelimiter &
                        _94_UNIQUE_NO & DataDelimiter &
                        _95_ROW_NO & DataDelimiter &
                        _96_I3B_UNIQUE_NO & DataDelimiter &
                        _98_ACCOUNT_NO & DataDelimiter &
                        _99_ACCOUNT_TYPE & DataDelimiter &
                        _901_SHARIA_INDICATOR & DataDelimiter &
                        _126_OPENING_BALANCE & DataDelimiter &
                        _127_ACCOUNT_START_DATE & DataDelimiter &
                        _128_CLOSING_BALANCE & DataDelimiter &
                        _129_ACCOUNT_CLOSING_DATE & DataDelimiter &
                        _80_NET_RETURN_ON_INVESTMENT_SOURCE_CODE & DataDelimiter &
                        _81_NET_RETURN_ON_INVESTMENT & DataDelimiter &
                        _82_INTEREST_SOURCE_CODE & DataDelimiter &
                        _83_INTEREST & DataDelimiter &
                        _84_DIVIDENDS_SOURCE_CODE & DataDelimiter &
                        _85_DIVIDENDS & DataDelimiter &
                        _86_CAPITAL_SOURCE_CODE & DataDelimiter &
                        _87_CAPITAL_GAIN & DataDelimiter &
                        _801_OTHER_SOURCE_CODE & DataDelimiter &
                        _802_OTHER & DataDelimiter &
                        _88_MARKET_VALUE & DataDelimiter &
                        _89_TRANSACTION_VALUE_INDICATOR

        'IncomeDetails = _91_SEC_ID & DataDelimiter & _
        '                _92_RECORD_TYPE & DataDelimiter & _
        '                _93_RECORD_STATUS & DataDelimiter & _
        '                _94_UNIQUE_NO & DataDelimiter & _
        '                _95_ROW_NO & DataDelimiter & _
        '                _96_I3B_UNIQUE_NO & DataDelimiter & _
        '                _97_INCOME_SOURCE_CODE & DataDelimiter & _
        '                _98_ACCOUNT_NO & DataDelimiter & _
        '                _99_ACCOUNT_TYPE & DataDelimiter & _
        '                _100_MARCH_CREDITS & DataDelimiter & _
        '                _101_APRIL_CREDITS & DataDelimiter & _
        '                _102_MAY_CREDITS & DataDelimiter & _
        '                _103_JUNE_CREDITS & DataDelimiter & _
        '                _104_JULY_CREDITS & DataDelimiter & _
        '                _105_AUGUST_CREDITS & DataDelimiter & _
        '                _106_SEPTEMBER_CREDITS & DataDelimiter & _
        '                _107_OCTOBER_CREDITS & DataDelimiter & _
        '                _108_NOVEMBER_CREDITS & DataDelimiter & _
        '                _109_DECEMBER_CREDITS & DataDelimiter & _
        '                _110_JANUARY_CREDITS & DataDelimiter & _
        '                _111_FEBRUARY_CREDITS & DataDelimiter & _
        '                _112_MARCH_DEBITS & DataDelimiter & _
        '                _113_APRIL_DEBITS & DataDelimiter & _
        '                _114_MAY_DEBITS & DataDelimiter & _
        '                _115_JUNE_DEBITS & DataDelimiter & _
        '                _116_JULY_DEBITS & DataDelimiter & _
        '                _117_AUGUST_DEBITS & DataDelimiter & _
        '                _118_SEPTEMBER_DEBITS & DataDelimiter & _
        '                _119_OCTOBER_DEBITS & DataDelimiter & _
        '                _120_NOVEMBER_DEBITS & DataDelimiter & _
        '                _121_DECEMBER_DEBITS & DataDelimiter & _
        '                _122_JANUARY_DEBITS & DataDelimiter & _
        '                _123_FEBRUARY_DEBITS & DataDelimiter & _
        '                _124_TOTAL_INTEREST_PAID & DataDelimiter & _
        '                _125_TOTAL_INTEREST_EARNED & DataDelimiter & _
        '                _126_OPENING_BALANCE & DataDelimiter & _
        '                _127_ACCOUNT_START_DATE & DataDelimiter & _
        '                _128_CLOSING_BALANCE & DataDelimiter & _
        '                _129_ACCOUNT_CLOSING_DATE & DataDelimiter & _
        '                _130_FOREIGN_TAX_PAID

    End Function

    Private Sub WriteTransactionDetails(ByVal AccountInfo As Object, ByVal AHFD_UniqueNo As String)

        Const _90_SEC_ID As String = "B"
        Const _91_RECORD_TYPE As String = "ATD"
        Dim SourceCodeContribution As String = "4219"
        Dim SourceCodeTransferIn As String = "4246"
        Dim SourceCodeTransferOut As String = "4247"
        Dim SourceCodeWithdrawal As String = "4248"
        Dim TransactionTypeCodeContribution As String = "01"
        Dim TransactionTypeCodeTransferIn As String = "02"
        Dim TransactionTypeCodeTransferOut As String = "03"
        Dim TransactionTypeCodeWithdrawal As String = "04"
        Dim ContributionCount As Integer = 0
        Dim TransferInCount As Integer = 0
        Dim TransferOutCount As Integer = 0
        Dim WithdrawalCount As Integer = 0
        Dim _92_RECORD_STATUS As String
        Dim _93_UNIQUE_NO As String = "Unique_No"  'TotalRecords.ToString
        Dim _94_ROW_NO As String = "Rec_No"
        'Dim _95_I3B_UNIQUE_NO As String = AccountInfo(AccountFields.taxCIF)    'FixedLengthString(AccountInfo(AccountFields.taxCIF), 25, "Left", " ")
        Dim _95_I3B_UNIQUE_NO As String = AHFD_UniqueNo.Replace("Unique_No", "AHFD_Unique_No")
        Dim _96_TRAN_NO As String
        Dim _97_TRAN_DATE As String
        Dim _98_TRAN_TYPE As String
        Dim _99_SOURCE_CODE As String
        Dim _100_TRAN_VALUE As String

        Dim TFSATransactionDetails As String
        Dim PeriodStart As String
        Dim PeriodEnd As String

        PeriodStart = Left(SARSForm.PeriodStart(SARSForm.txtPeriodLength.Text), 6) & "01"
        PeriodEnd = Left(SARSForm.PeriodEnd, 6) & "01"

        '21/10/2016:  We will only include deposits at this time
        '27/09/2021:  Now we have to include at least one of each type of transaction, it would seem.  SARS double-speak in play.......
        Dim strSQLTrans As String = "DECLARE	@Trans AS TABLE	(trnId INT
						                                        ,trnTransactionDate VARCHAR(8)
						                                        ,pcCode VARCHAR(8)
						                                        ,pcDrCr VARCHAR(5)
						                                        ,trnAmt NUMERIC(18,2)
						                                        ,IsCorrection BIT
						                                        ,WasCorrected BIT)

                                        INSERT INTO @Trans

	                                        SELECT	 tblTransactions.trnId
			                                        ,tblTransactions.trnTransactionDate
			                                        ,tblTransactions.pcCode
			                                        ,mblPostingCodes.pcDrCr
			                                        ,tblTransactions.trnAmt
			                                        ,tblTransactions.trnCorrection
			                                        ,0
	                                        FROM   tblAccount INNER JOIN
			                                        tblTransactions ON tblAccount.accId = tblTransactions.accId INNER JOIN
			                                        mblPostingCodes ON tblTransactions.pcCode = mblPostingCodes.pcCode
	                                        WHERE  tblAccount.accNO = '" & AccountInfo(AccountFields.taxAccountNo) & "'
				                                        AND mblPostingCodes.pcIsInterest = 0
				                                        AND mblPostingCodes.pcDrCr IN ('Dr', 'Cr')
				                                        AND tblTransactions.pcCode NOT IN ('F6', 'RC', '16', '17', '68', '69', 'TFF6')
				                                        AND tblTransactions.trnTransBatchMonth >= '" & PeriodStart & "'
				                                        AND tblTransactions.trnTransBatchMonth <= '" & PeriodEnd & "'
				                                        AND tblTransactions.Deleted = 0
	                                        ORDER BY   tblTransactions.trnId

                                        UPDATE	Trans
                                        SET		WasCorrected = 1
                                        FROM	@Trans AS Trans INNER JOIN
		                                        @Trans AS Corrections ON Trans.trnTransactionDate = Corrections.trnTransactionDate
									                                        AND Trans.pcCode = Corrections.pcCode
									                                        AND Trans.trnAmt + Corrections.trnAmt = 0
									                                        AND Corrections.IsCorrection = 1

                                        SELECT	 trnId
		                                        ,trnTransactionDate
		                                        ,pcCode
		                                        ,pcDrCr
		                                        ,trnAmt
                                        FROM	@Trans
                                        WHERE	IsCorrection = 0
			                                        AND WasCorrected = 0"

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As New SqlCommand(strSQLTrans, cn)

        Try
            cn.Open()
            drSQL = cmd.ExecuteReader()
            While drSQL.Read()

                _96_TRAN_NO = drSQL.Item(0).ToString
                _97_TRAN_DATE = DateWithDashes(drSQL.Item(1))

                'At the moment, we have no real way of identifying transfers in or out externally, so we ignore those.
                'Transfers internally are also suspect, as human error and posting code issues make identifying them reliably difficult, so we ignore these as well.
                'See Jira Jobcard SUP-755:  We are now forcing users to make use of proper procedures and well-defined posting codes, so we
                '                           have a better chance of correctly identifying transaction types
                Select Case drSQL.Item(2)
                    Case "TFC", "TFC6"
                        _98_TRAN_TYPE = TransactionTypeCodeContribution
                        _99_SOURCE_CODE = SourceCodeContribution
                        ContributionCount += 1
                    Case "TFTI", "TF86"
                        _98_TRAN_TYPE = TransactionTypeCodeTransferIn
                        _99_SOURCE_CODE = SourceCodeTransferIn
                        TransferInCount += 1
                    Case "TFTO", "TFCT", "TF66"
                        _98_TRAN_TYPE = TransactionTypeCodeTransferOut
                        _99_SOURCE_CODE = SourceCodeTransferOut
                        TransferOutCount += 1
                    Case "TFW", "TFCW", "TF76", "TFW6"
                        _98_TRAN_TYPE = TransactionTypeCodeWithdrawal
                        _99_SOURCE_CODE = SourceCodeWithdrawal
                        WithdrawalCount += 1
                    Case Else
                        Select Case drSQL.Item(3)
                            Case "Cr"
                                _98_TRAN_TYPE = TransactionTypeCodeContribution
                                _99_SOURCE_CODE = SourceCodeContribution
                                ContributionCount += 1
                            Case "Dr"
                                _98_TRAN_TYPE = TransactionTypeCodeWithdrawal
                                _99_SOURCE_CODE = SourceCodeWithdrawal
                                WithdrawalCount += 1
                        End Select
                End Select

                _100_TRAN_VALUE = FormatNumber(Math.Abs(drSQL.Item(4)), 2, TriState.True, TriState.False, TriState.False)
                TotalMoney = TotalMoney + _100_TRAN_VALUE

                '_93_UNIQUE_NO = GetExistingUniqueNo(_95_I3B_UNIQUE_NO, _91_RECORD_TYPE, SARSForm.TestFile, _96_TRAN_NO)
                _93_UNIQUE_NO = GetExistingUniqueNo(AccountInfo(AccountFields.taxCIF), _91_RECORD_TYPE, SARSForm.TestFile, _96_TRAN_NO)
                If _93_UNIQUE_NO <> "Unique_No" Then
                    '_95_ROW_NO = _94_UNIQUE_NO
                    _92_RECORD_STATUS = "C"
                Else
                    _92_RECORD_STATUS = "N"
                End If

                TFSATransactionDetails = _90_SEC_ID & DataDelimiter &
                                         _91_RECORD_TYPE & DataDelimiter &
                                         _92_RECORD_STATUS & DataDelimiter &
                                         _93_UNIQUE_NO & DataDelimiter &
                                         _94_ROW_NO & DataDelimiter &
                                         _95_I3B_UNIQUE_NO & DataDelimiter &
                                         _96_TRAN_NO & DataDelimiter &
                                         _97_TRAN_DATE & DataDelimiter &
                                         _98_TRAN_TYPE & DataDelimiter &
                                         _99_SOURCE_CODE & DataDelimiter &
                                         _100_TRAN_VALUE

                WriteToDatabase(IT3sClientDataTable, TFSATransactionDetails)  'True

            End While
            cn.Close()
            drSQL = Nothing
        Catch ex As Exception
        End Try

        'Now we create zero records for any category of transaction which has had no entries
        If ContributionCount = 0 Then

            _93_UNIQUE_NO = GetExistingUniqueNo(AccountInfo(AccountFields.taxCIF), _91_RECORD_TYPE, SARSForm.TestFile, SourceCode:=SourceCodeContribution)
            If _93_UNIQUE_NO <> "Unique_No" Then
                '_95_ROW_NO = _94_UNIQUE_NO
                _92_RECORD_STATUS = "C"
            Else
                _92_RECORD_STATUS = "N"
            End If

            TFSATransactionDetails = _90_SEC_ID & DataDelimiter &
                                         _91_RECORD_TYPE & DataDelimiter &
                                         _92_RECORD_STATUS & DataDelimiter &
                                         _93_UNIQUE_NO & DataDelimiter &
                                         _94_ROW_NO & DataDelimiter &
                                         _95_I3B_UNIQUE_NO & DataDelimiter &
                                         "<" & AccountInfo(AccountFields.taxAccountNo) & ">" & DataDelimiter &
                                         DateWithDashes(SARSForm.PeriodEnd) & DataDelimiter &
                                         "01" & DataDelimiter &
                                         SourceCodeContribution & DataDelimiter &
                                         "0.00"

            WriteToDatabase(IT3sClientDataTable, TFSATransactionDetails)  'True

        End If

        If TransferInCount = 0 Then

            _93_UNIQUE_NO = GetExistingUniqueNo(AccountInfo(AccountFields.taxCIF), _91_RECORD_TYPE, SARSForm.TestFile, SourceCode:=SourceCodeTransferIn)
            If _93_UNIQUE_NO <> "Unique_No" Then
                '_95_ROW_NO = _94_UNIQUE_NO
                _92_RECORD_STATUS = "C"
            Else
                _92_RECORD_STATUS = "N"
            End If

            TFSATransactionDetails = _90_SEC_ID & DataDelimiter &
                                        _91_RECORD_TYPE & DataDelimiter &
                                        _92_RECORD_STATUS & DataDelimiter &
                                        _93_UNIQUE_NO & DataDelimiter &
                                        _94_ROW_NO & DataDelimiter &
                                        _95_I3B_UNIQUE_NO & DataDelimiter &
                                        "<" & AccountInfo(AccountFields.taxAccountNo) & ">" & DataDelimiter &
                                        DateWithDashes(SARSForm.PeriodEnd) & DataDelimiter &
                                        "02" & DataDelimiter &
                                        SourceCodeTransferIn & DataDelimiter &
                                        "0.00"

            WriteToDatabase(IT3sClientDataTable, TFSATransactionDetails)  'True

        End If

        If TransferOutCount = 0 Then

            _93_UNIQUE_NO = GetExistingUniqueNo(AccountInfo(AccountFields.taxCIF), _91_RECORD_TYPE, SARSForm.TestFile, SourceCode:=SourceCodeTransferOut)
            If _93_UNIQUE_NO <> "Unique_No" Then
                '_95_ROW_NO = _94_UNIQUE_NO
                _92_RECORD_STATUS = "C"
            Else
                _92_RECORD_STATUS = "N"
            End If

            TFSATransactionDetails = _90_SEC_ID & DataDelimiter &
                                        _91_RECORD_TYPE & DataDelimiter &
                                        _92_RECORD_STATUS & DataDelimiter &
                                        _93_UNIQUE_NO & DataDelimiter &
                                        _94_ROW_NO & DataDelimiter &
                                        _95_I3B_UNIQUE_NO & DataDelimiter &
                                        "<" & AccountInfo(AccountFields.taxAccountNo) & ">" & DataDelimiter &
                                        DateWithDashes(SARSForm.PeriodEnd) & DataDelimiter &
                                        "03" & DataDelimiter &
                                        SourceCodeTransferOut & DataDelimiter &
                                        "0.00"

            WriteToDatabase(IT3sClientDataTable, TFSATransactionDetails)  'True

        End If

        If WithdrawalCount = 0 Then

            _93_UNIQUE_NO = GetExistingUniqueNo(AccountInfo(AccountFields.taxCIF), _91_RECORD_TYPE, SARSForm.TestFile, SourceCode:=SourceCodeWithdrawal)
            If _93_UNIQUE_NO <> "Unique_No" Then
                '_95_ROW_NO = _94_UNIQUE_NO
                _92_RECORD_STATUS = "C"
            Else
                _92_RECORD_STATUS = "N"
            End If

            TFSATransactionDetails = _90_SEC_ID & DataDelimiter &
                                        _91_RECORD_TYPE & DataDelimiter &
                                        _92_RECORD_STATUS & DataDelimiter &
                                        _93_UNIQUE_NO & DataDelimiter &
                                        _94_ROW_NO & DataDelimiter &
                                        _95_I3B_UNIQUE_NO & DataDelimiter &
                                        "<" & AccountInfo(AccountFields.taxAccountNo) & ">" & DataDelimiter &
                                        DateWithDashes(SARSForm.PeriodEnd) & DataDelimiter &
                                        "04" & DataDelimiter &
                                        SourceCodeWithdrawal & DataDelimiter &
                                        "0.00"

            WriteToDatabase(IT3sClientDataTable, TFSATransactionDetails)  'True

        End If

    End Sub

    'Private Sub WriteTransactionDetails(ByVal AccountInfo As Object, ByVal AHFD_UniqueNo As String)

    '    Const _90_SEC_ID As String = "B"
    '    Const _91_RECORD_TYPE As String = "ATD"
    '    Dim SourceCodeContribution As String = "4219"
    '    Dim SourceCodeTransferIn As String = "4246"
    '    Dim SourceCodeTransferOut As String = "4247"
    '    Dim SourceCodeWithdrawal As String = "4248"
    '    Dim ContributionCount As Integer = 0
    '    Dim TransferInCount As Integer = 0
    '    Dim TransferOutCount As Integer = 0
    '    Dim WithdrawalCount As Integer = 0
    '    Dim _92_RECORD_STATUS As String
    '    Dim _93_UNIQUE_NO As String = "Unique_No"  'TotalRecords.ToString
    '    Dim _94_ROW_NO As String = "Rec_No"
    '    'Dim _95_I3B_UNIQUE_NO As String = AccountInfo(AccountFields.taxCIF)    'FixedLengthString(AccountInfo(AccountFields.taxCIF), 25, "Left", " ")
    '    Dim _95_I3B_UNIQUE_NO As String = AHFD_UniqueNo.Replace("Unique_No", "AHFD_Unique_No")
    '    Dim _96_TRAN_NO As String
    '    Dim _97_TRAN_DATE As String
    '    Dim _98_TRAN_TYPE As String
    '    Dim _99_SOURCE_CODE As String
    '    Dim _100_TRAN_VALUE As String

    '    Dim TFSATransactionDetails As String
    '    Dim PeriodStart As String
    '    Dim PeriodEnd As String

    '    PeriodStart = Left(SARSForm.PeriodStart(SARSForm.txtPeriodLength.Text), 6) & "01"
    '    PeriodEnd = Left(SARSForm.PeriodEnd, 6) & "01"

    '    '21/10/2016:  We will only include deposits at this time
    '    '27/09/2021:  Now we have to include at least one of each type of transaction, it would seem.  SARS double-speak in play.......
    '    Dim strSQLTrans As String = "Select tblTransactions.trnId, " &
    '                                       "tblTransactions.trnTransactionDate, " &
    '                                       "tblTransactions.pcCode, " &
    '                                       "mblPostingCodes.pcDrCr, " &
    '                                       "tblTransactions.trnAmt, " &
    '                                       "tblTransactions.trnCorrection " &
    '                                "FROM   tblAccount INNER JOIN " &
    '                                       "tblTransactions On tblAccount.accId = tblTransactions.accId INNER JOIN " &
    '                                       "mblPostingCodes On tblTransactions.pcCode = mblPostingCodes.pcCode " &
    '                                "WHERE  tblAccount.accNO = '" & AccountInfo(AccountFields.taxAccountNo) & "' " &
    '                                            "AND mblPostingCodes.pcIsInterest = 0 " &
    '                                            "AND mblPostingCodes.pcDrCr IN ('Dr', 'Cr') " &
    '                                            "AND tblTransactions.pcCode NOT IN ('F6', 'RC', '16', '17', '68', '69') " &
    '                                            "AND tblTransactions.trnTransBatchMonth >= '" & PeriodStart & "' " &
    '                                            "AND tblTransactions.trnTransBatchMonth <= '" & PeriodEnd & "' " &
    '                                            "/*AND tblTransactions.trnAmt > 0 */" &
    '                                            "AND tblTransactions.Deleted = 0 " &
    '                                "ORDER BY   tblTransactions.trnId"

    '    Dim cn As SqlConnection = New SqlConnection(strConnection)
    '    Dim drSQL As SqlDataReader
    '    Dim cmd As New SqlCommand(strSQLTrans, cn)

    '    Try
    '        cn.Open()
    '        drSQL = cmd.ExecuteReader()
    '        While drSQL.Read()

    '            _96_TRAN_NO = drSQL.Item(0).ToString
    '            _97_TRAN_DATE = DateWithDashes(drSQL.Item(1))

    '            'At the moment, we have no real way of identifying transfers in or out externally, so we ignore those.
    '            'Transfers internally are also suspect, as human error and posting code issues make identifying them reliably difficult, so we ignore these as well.
    '            Select Case drSQL.Item(3)
    '                Case "Cr"
    '                    If drSQL.Item(5) Then
    '                        _98_TRAN_TYPE = "04"
    '                        _99_SOURCE_CODE = SourceCodeWithdrawal
    '                        WithdrawalCount += 1
    '                    Else
    '                        _98_TRAN_TYPE = "01"
    '                        _99_SOURCE_CODE = SourceCodeContribution
    '                        ContributionCount += 1
    '                    End If
    '                Case "Dr"
    '                    If drSQL.Item(5) Then
    '                        _98_TRAN_TYPE = "01"
    '                        _99_SOURCE_CODE = SourceCodeContribution
    '                        ContributionCount += 1
    '                    Else
    '                        _98_TRAN_TYPE = "04"
    '                        _99_SOURCE_CODE = SourceCodeWithdrawal
    '                        WithdrawalCount += 1
    '                    End If
    '            End Select

    '            _100_TRAN_VALUE = FormatNumber(Math.Abs(drSQL.Item(4)), 2, TriState.True, TriState.False, TriState.False)
    '            TotalMoney = TotalMoney + _100_TRAN_VALUE

    '            '_93_UNIQUE_NO = GetExistingUniqueNo(_95_I3B_UNIQUE_NO, _91_RECORD_TYPE, SARSForm.TestFile, _96_TRAN_NO)
    '            _93_UNIQUE_NO = GetExistingUniqueNo(AccountInfo(AccountFields.taxCIF), _91_RECORD_TYPE, SARSForm.TestFile, _96_TRAN_NO)
    '            If _93_UNIQUE_NO <> "Unique_No" Then
    '                '_95_ROW_NO = _94_UNIQUE_NO
    '                _92_RECORD_STATUS = "C"
    '            Else
    '                _92_RECORD_STATUS = "N"
    '            End If

    '            TFSATransactionDetails = _90_SEC_ID & DataDelimiter &
    '                                     _91_RECORD_TYPE & DataDelimiter &
    '                                     _92_RECORD_STATUS & DataDelimiter &
    '                                     _93_UNIQUE_NO & DataDelimiter &
    '                                     _94_ROW_NO & DataDelimiter &
    '                                     _95_I3B_UNIQUE_NO & DataDelimiter &
    '                                     _96_TRAN_NO & DataDelimiter &
    '                                     _97_TRAN_DATE & DataDelimiter &
    '                                     _98_TRAN_TYPE & DataDelimiter &
    '                                     _99_SOURCE_CODE & DataDelimiter &
    '                                     _100_TRAN_VALUE

    '            WriteToDatabase(IT3sClientDataTable, TFSATransactionDetails)  'True

    '        End While
    '        cn.Close()
    '        drSQL = Nothing
    '    Catch ex As Exception
    '    End Try

    '    'Now we create zero records for any category of transaction which has had no entries
    '    If ContributionCount = 0 Then

    '        _93_UNIQUE_NO = GetExistingUniqueNo(AccountInfo(AccountFields.taxCIF), _91_RECORD_TYPE, SARSForm.TestFile, SourceCode:=SourceCodeContribution)
    '        If _93_UNIQUE_NO <> "Unique_No" Then
    '            '_95_ROW_NO = _94_UNIQUE_NO
    '            _92_RECORD_STATUS = "C"
    '        Else
    '            _92_RECORD_STATUS = "N"
    '        End If

    '        TFSATransactionDetails = _90_SEC_ID & DataDelimiter &
    '                                     _91_RECORD_TYPE & DataDelimiter &
    '                                     _92_RECORD_STATUS & DataDelimiter &
    '                                     _93_UNIQUE_NO & DataDelimiter &
    '                                     _94_ROW_NO & DataDelimiter &
    '                                     _95_I3B_UNIQUE_NO & DataDelimiter &
    '                                     "<" & AccountInfo(AccountFields.taxAccountNo) & ">" & DataDelimiter &
    '                                     DateWithDashes(SARSForm.PeriodEnd) & DataDelimiter &
    '                                     "01" & DataDelimiter &
    '                                     SourceCodeContribution & DataDelimiter &
    '                                     "0.00"

    '        WriteToDatabase(IT3sClientDataTable, TFSATransactionDetails)  'True

    '    End If

    '    If TransferInCount = 0 Then

    '        _93_UNIQUE_NO = GetExistingUniqueNo(AccountInfo(AccountFields.taxCIF), _91_RECORD_TYPE, SARSForm.TestFile, SourceCode:=SourceCodeTransferIn)
    '        If _93_UNIQUE_NO <> "Unique_No" Then
    '            '_95_ROW_NO = _94_UNIQUE_NO
    '            _92_RECORD_STATUS = "C"
    '        Else
    '            _92_RECORD_STATUS = "N"
    '        End If

    '        TFSATransactionDetails = _90_SEC_ID & DataDelimiter &
    '                                    _91_RECORD_TYPE & DataDelimiter &
    '                                    _92_RECORD_STATUS & DataDelimiter &
    '                                    _93_UNIQUE_NO & DataDelimiter &
    '                                    _94_ROW_NO & DataDelimiter &
    '                                    _95_I3B_UNIQUE_NO & DataDelimiter &
    '                                    "<" & AccountInfo(AccountFields.taxAccountNo) & ">" & DataDelimiter &
    '                                    DateWithDashes(SARSForm.PeriodEnd) & DataDelimiter &
    '                                    "02" & DataDelimiter &
    '                                    SourceCodeTransferIn & DataDelimiter &
    '                                    "0.00"

    '        WriteToDatabase(IT3sClientDataTable, TFSATransactionDetails)  'True

    '    End If

    '    If TransferOutCount = 0 Then

    '        _93_UNIQUE_NO = GetExistingUniqueNo(AccountInfo(AccountFields.taxCIF), _91_RECORD_TYPE, SARSForm.TestFile, SourceCode:=SourceCodeTransferOut)
    '        If _93_UNIQUE_NO <> "Unique_No" Then
    '            '_95_ROW_NO = _94_UNIQUE_NO
    '            _92_RECORD_STATUS = "C"
    '        Else
    '            _92_RECORD_STATUS = "N"
    '        End If

    '        TFSATransactionDetails = _90_SEC_ID & DataDelimiter &
    '                                    _91_RECORD_TYPE & DataDelimiter &
    '                                    _92_RECORD_STATUS & DataDelimiter &
    '                                    _93_UNIQUE_NO & DataDelimiter &
    '                                    _94_ROW_NO & DataDelimiter &
    '                                    _95_I3B_UNIQUE_NO & DataDelimiter &
    '                                    "<" & AccountInfo(AccountFields.taxAccountNo) & ">" & DataDelimiter &
    '                                    DateWithDashes(SARSForm.PeriodEnd) & DataDelimiter &
    '                                    "03" & DataDelimiter &
    '                                    SourceCodeTransferOut & DataDelimiter &
    '                                    "0.00"

    '        WriteToDatabase(IT3sClientDataTable, TFSATransactionDetails)  'True

    '    End If

    '    If WithdrawalCount = 0 Then

    '        _93_UNIQUE_NO = GetExistingUniqueNo(AccountInfo(AccountFields.taxCIF), _91_RECORD_TYPE, SARSForm.TestFile, SourceCode:=SourceCodeWithdrawal)
    '        If _93_UNIQUE_NO <> "Unique_No" Then
    '            '_95_ROW_NO = _94_UNIQUE_NO
    '            _92_RECORD_STATUS = "C"
    '        Else
    '            _92_RECORD_STATUS = "N"
    '        End If

    '        TFSATransactionDetails = _90_SEC_ID & DataDelimiter &
    '                                    _91_RECORD_TYPE & DataDelimiter &
    '                                    _92_RECORD_STATUS & DataDelimiter &
    '                                    _93_UNIQUE_NO & DataDelimiter &
    '                                    _94_ROW_NO & DataDelimiter &
    '                                    _95_I3B_UNIQUE_NO & DataDelimiter &
    '                                    "<" & AccountInfo(AccountFields.taxAccountNo) & ">" & DataDelimiter &
    '                                    DateWithDashes(SARSForm.PeriodEnd) & DataDelimiter &
    '                                    "04" & DataDelimiter &
    '                                    SourceCodeWithdrawal & DataDelimiter &
    '                                    "0.00"

    '        WriteToDatabase(IT3sClientDataTable, TFSATransactionDetails)  'True

    '    End If

    'End Sub

    'Private Sub WriteTransactionDetails(ByVal AccountInfo As Object)

    '    Const _90_SEC_ID As String = "B"
    '    Const _91_RECORD_TYPE As String = "ATD"
    '    Dim _92_RECORD_STATUS As String
    '    Dim _93_UNIQUE_NO As String = "Unique_No"  'TotalRecords.ToString
    '    Dim _94_ROW_NO As String = "Rec_No"
    '    Dim _95_I3B_UNIQUE_NO As String = AccountInfo(AccountFields.taxCIF)    'FixedLengthString(AccountInfo(AccountFields.taxCIF), 25, "Left", " ")
    '    Dim _96_TRAN_NO As String
    '    Dim _97_TRAN_DATE As String
    '    Dim _98_TRAN_TYPE As String
    '    Dim _99_SOURCE_CODE As String = "4219"
    '    Dim _100_TRAN_VALUE As String

    '    Dim TFSATransactionDetails As String
    '    Dim PeriodStart As String
    '    Dim PeriodEnd As String

    '    PeriodStart = Left(SARSForm.PeriodStart(SARSForm.txtPeriodLength.Text), 6) & "01"
    '    PeriodEnd = Left(SARSForm.PeriodEnd, 6) & "01"

    '    'Dim strSQLTrans As String = "SELECT tblTransactions.trnId, " & _
    '    '                                   "tblTransactions.trnTransactionDate, " & _
    '    '                                   "tblTransactions.pcCode, " & _
    '    '                                   "mblPostingCodes.pcDrCr, " & _
    '    '                                   "tblTransactions.trnAmt, " & _
    '    '                                   "tblTransactions.trnCorrection " & _
    '    '                            "FROM   tblAccount INNER JOIN " & _
    '    '                                   "tblTransactions ON tblAccount.accId = tblTransactions.accId INNER JOIN " & _
    '    '                                   "mblPostingCodes ON tblTransactions.pcCode = mblPostingCodes.pcCode " & _
    '    '                            "WHERE  tblAccount.accNO = '" & AccountInfo(AccountFields.taxAccountNo) & "' " & _
    '    '                                        "AND mblPostingCodes.pcIsInterest = 0 " & _
    '    '                                        "AND tblTransactions.pcCode NOT IN ('F6', 'RC', '16', '17', '68', '69') " & _
    '    '                                        "AND tblTransactions.trnTransBatchMonth >= '" & PeriodStart & "' " & _
    '    '                                        "AND tblTransactions.trnTransBatchMonth <= '" & PeriodEnd & "' " & _
    '    '                                        "AND tblTransactions.trnAmt <> 0 " & _
    '    '                                        "AND tblTransactions.Deleted = 0 " & _
    '    '                            "ORDER BY   tblTransactions.trnId"

    '    '21/10/2016:  We will only include deposits at this time
    '    Dim strSQLTrans As String = "SELECT tblTransactions.trnId, " &
    '                                       "tblTransactions.trnTransactionDate, " &
    '                                       "tblTransactions.pcCode, " &
    '                                       "mblPostingCodes.pcDrCr, " &
    '                                       "tblTransactions.trnAmt, " &
    '                                       "tblTransactions.trnCorrection " &
    '                                "FROM   tblAccount INNER JOIN " &
    '                                       "tblTransactions ON tblAccount.accId = tblTransactions.accId INNER JOIN " &
    '                                       "mblPostingCodes ON tblTransactions.pcCode = mblPostingCodes.pcCode " &
    '                                "WHERE  tblAccount.accNO = '" & AccountInfo(AccountFields.taxAccountNo) & "' " &
    '                                            "AND mblPostingCodes.pcIsInterest = 0 " &
    '                                            "AND mblPostingCodes.pcDrCr = 'Cr' " &
    '                                            "AND tblTransactions.pcCode NOT IN ('F6', 'RC', '16', '17', '68', '69') " &
    '                                            "AND tblTransactions.trnTransBatchMonth >= '" & PeriodStart & "' " &
    '                                            "AND tblTransactions.trnTransBatchMonth <= '" & PeriodEnd & "' " &
    '                                            "AND tblTransactions.trnAmt > 0 " &
    '                                            "AND tblTransactions.Deleted = 0 " &
    '                                "ORDER BY   tblTransactions.trnId"

    '    Dim cn As SqlConnection = New SqlConnection(strConnection)
    '    Dim drSQL As SqlDataReader
    '    Dim cmd As New SqlCommand(strSQLTrans, cn)

    '    Try
    '        cn.Open()
    '        drSQL = cmd.ExecuteReader()
    '        While drSQL.Read()

    '            _96_TRAN_NO = drSQL.Item(0).ToString
    '            _97_TRAN_DATE = DateWithDashes(drSQL.Item(1))

    '            'Currently not sure of the definition of the SARS "Transfer In" and "Transfer Out" transaction types,
    '            'and they are in any case only valid from 01/11/2016  (according to the reponse files, even thought the BRS says 01/03/2016),
    '            'so we will treat everything as either a "Contribution" (code 01) or a "Withdrawal" (code 04) for now
    '            'Select Case drSQL.Item(2)
    '            '    Case "80"
    '            '        If drSQL.Item(4) Then
    '            '            _98_TRAN_TYPE = "04"
    '            '        Else
    '            '            _98_TRAN_TYPE = "01"
    '            '        End If
    '            '    Case "86"
    '            '        If drSQL.Item(4) Then
    '            '            _98_TRAN_TYPE = "03"
    '            '        Else
    '            '            _98_TRAN_TYPE = "02"
    '            '        End If
    '            '    Case "06", "66", "63", "64"
    '            '        If drSQL.Item(4) Then
    '            '            _98_TRAN_TYPE = "02"
    '            '        Else
    '            '            _98_TRAN_TYPE = "03"
    '            '        End If
    '            '    Case "00", "67", "70", "74"
    '            '        If drSQL.Item(4) Then
    '            '            _98_TRAN_TYPE = "01"
    '            '        Else
    '            '            _98_TRAN_TYPE = "04"
    '            '        End If
    '            '    Case Else
    '            '        If drSQL.Item(3) > 0 Then
    '            '            _98_TRAN_TYPE = "02"
    '            '        Else
    '            '            _98_TRAN_TYPE = "03"
    '            '        End If
    '            'End Select

    '            'tblTransactions.trnId, " & _"
    '            '                           "tblTransactions.trnTransactionDate, " & _
    '            '                           "tblTransactions.pcCode, " & _
    '            '                           "mblPostingCodes.pcDrCr, " & _
    '            '                           "tblTransactions.trnAmt, " & _
    '            '                           "tblTransactions.trnCorrection " & 

    '            Select Case drSQL.Item(3)
    '                Case "Cr"
    '                    If drSQL.Item(5) Then
    '                        _98_TRAN_TYPE = "04"
    '                    Else
    '                        _98_TRAN_TYPE = "01"
    '                    End If
    '                Case "Dr"
    '                    If drSQL.Item(5) Then
    '                        _98_TRAN_TYPE = "01"
    '                    Else
    '                        _98_TRAN_TYPE = "04"
    '                    End If
    '            End Select

    '            _100_TRAN_VALUE = FormatNumber(Math.Abs(drSQL.Item(4)), 2, TriState.True, TriState.False, TriState.False)
    '            'TotalMoney = TotalMoney + _100_TRAN_VALUE

    '            _93_UNIQUE_NO = GetExistingUniqueNo(_95_I3B_UNIQUE_NO, _91_RECORD_TYPE, SARSForm.TestFile, _96_TRAN_NO)
    '            If _93_UNIQUE_NO <> "Unique_No" Then
    '                '_95_ROW_NO = _94_UNIQUE_NO
    '                _92_RECORD_STATUS = "C"
    '            Else
    '                _92_RECORD_STATUS = "N"
    '            End If

    '            TFSATransactionDetails = _90_SEC_ID & DataDelimiter &
    '                                     _91_RECORD_TYPE & DataDelimiter &
    '                                     _92_RECORD_STATUS & DataDelimiter &
    '                                     _93_UNIQUE_NO & DataDelimiter &
    '                                     _94_ROW_NO & DataDelimiter &
    '                                     _95_I3B_UNIQUE_NO & DataDelimiter &
    '                                     _96_TRAN_NO & DataDelimiter &
    '                                     _97_TRAN_DATE & DataDelimiter &
    '                                     _98_TRAN_TYPE & DataDelimiter &
    '                                     _99_SOURCE_CODE & DataDelimiter &
    '                                     _100_TRAN_VALUE

    '            WriteToDatabase(IT3sClientDataTable, TFSATransactionDetails)  'True

    '        End While
    '        cn.Close()
    '        drSQL = Nothing
    '    Catch ex As Exception
    '    End Try

    'End Sub

    'Private Function IncomeDetails(ByVal AccountInfo As Object) As String

    '    Const SEC_ID As String = "I"
    '    Const INCOME_NATURE As String = "4201"
    '    Const BRANCH_CODE As String = "000000"

    '    Dim IT3_PERS_ID As String
    '    Dim INCOME_PAID As String
    '    Dim ACCOUNT_NO As String
    '    Dim ACCOUNT_TYPE As String
    '    Dim START_DATE As String
    '    Dim START_BAL As String
    '    Dim START_BAL_SIGN As String
    '    Dim END_DATE As String
    '    Dim END_BAL As String
    '    Dim END_BAL_SIGN As String
    '    Dim FOREIGN_TAX_PAID As String

    '    Dim OpeningBalance As Integer
    '    Dim ClosingBalance As Integer
    '    Dim CloseDate As String
    '    Dim Interest As Integer

    '    IT3_PERS_ID = FixedLengthString(AccountInfo(0), 25, "Left", " ")
    '    Interest = CInt(FixNullNum((AccountInfo(1)) + FixNullNum(AccountInfo(2))) * 100)
    '    If Interest < 0 Then
    '        Interest = 0
    '    End If
    '    INCOME_PAID = FixedLengthString(Interest, 15, "Right", " ")
    '    ACCOUNT_NO = FixedLengthString(AccountInfo(3), 20, "Left", " ")
    '    Select Case AccountInfo(4)
    '        Case 11, 12, 14, 16
    '            ACCOUNT_TYPE = "03"
    '        Case 13, 18, 19
    '            ACCOUNT_TYPE = "06"
    '        Case 51
    '            ACCOUNT_TYPE = "05"
    '        Case Else
    '            ACCOUNT_TYPE = "17"
    '    End Select

    '    If AccountInfo(5) > SARSForm.PeriodStart(SARSForm.txtPeriodLength.Text)(SARSForm.txtPeriodLength.Text) Then
    '        START_DATE = AccountInfo(5)
    '    Else
    '        START_DATE = SARSForm.PeriodStart(SARSForm.txtPeriodLength.Text)
    '    End If
    '    OpeningBalance = AccountOpeningBalance(AccountInfo(7), START_DATE)
    '    START_BAL = FixedLengthString(OpeningBalance, 15, "Right", " ")
    '    If OpeningBalance >= 0 Then
    '        START_BAL_SIGN = "C"
    '    Else
    '        START_BAL_SIGN = "D"
    '    End If
    '    If AccountInfo(6) > 0 Then
    '        CloseDate = AccountCloseDate(AccountInfo(7))
    '    Else
    '        CloseDate = "0"
    '    End If
    '    If CloseDate > 0 And CloseDate < SARSForm.PeriodEnd Then
    '        END_DATE = CloseDate
    '    Else
    '        END_DATE = SARSForm.PeriodEnd
    '    End If
    '    ClosingBalance = AccountClosingBalance(AccountInfo(7), END_DATE)
    '    END_BAL = FixedLengthString(ClosingBalance, 15, "Right", " ")
    '    If ClosingBalance >= 0 Then
    '        END_BAL_SIGN = "C"
    '    Else
    '        END_BAL_SIGN = "D"
    '    End If
    '    FOREIGN_TAX_PAID = FixedLengthString(" ", 15, "Right", " ")

    '    If START_BAL = 0 Or INCOME_PAID = 0 Then
    '        IncomeDetails = ""
    '    Else
    '        IncomeDetails = SEC_ID & _
    '                        IT3_PERS_ID & _
    '                        INCOME_NATURE & _
    '                        INCOME_PAID & _
    '                        ACCOUNT_NO & _
    '                        BRANCH_CODE & _
    '                        ACCOUNT_TYPE & _
    '                        START_DATE & _
    '                        START_BAL & _
    '                        START_BAL_SIGN & _
    '                        END_DATE & _
    '                        END_BAL & _
    '                        END_BAL_SIGN & _
    '                        FOREIGN_TAX_PAID
    '    End If

    'End Function

    Private Function AccountTransactions(ByVal accNo As String) As Array

        Dim GrossTrans(12, 2) As String
        Dim PeriodStart As String
        Dim PeriodEnd As String
        Dim i As Byte

        For i = 0 To 12
            GrossTrans(i, 1) = "0.00"
            GrossTrans(i, 2) = "0.00"
        Next

        'If accNo = "008531110010" Then
        '    Debug.Write(" ")
        'End If

        PeriodStart = Left(SARSForm.PeriodStart(SARSForm.txtPeriodLength.Text), 6) & "01"
        PeriodEnd = Left(SARSForm.PeriodEnd, 6) & "01"

        Dim strSQLDebit As String = "SELECT trnTransBatchMonth, SUM(trnAmt) AS GrossDebits " &
                                    "FROM   tblAccount INNER JOIN " &
                                           "tblTransactions ON tblAccount.accId = tblTransactions.accId " &
                                    "WHERE  accNO = '" & accNo & "' AND " &
                                           "trnTransBatchMonth >= '" & PeriodStart & "' AND " &
                                           "trnTransBatchMonth <= '" & PeriodEnd & "' AND " &
                                           "trnAmt < 0 AND " &
                                           "tblTransactions.Deleted = 0 " &
                                    "GROUP BY trnTransBatchMonth"

        Dim strSQLCredit As String = "SELECT trnTransBatchMonth, SUM(trnAmt) AS GrossCredits " &
                                     "FROM   tblAccount INNER JOIN " &
                                            "tblTransactions ON tblAccount.accId = tblTransactions.accId " &
                                     "WHERE  accNO = '" & accNo & "' AND " &
                                            "trnTransBatchMonth >= '" & PeriodStart & "' AND " &
                                            "trnTransBatchMonth <= '" & PeriodEnd & "' AND " &
                                            "trnAmt > 0 AND " &
                                            "tblTransactions.Deleted = 0 " &
                                     "GROUP BY trnTransBatchMonth"

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmdDr As SqlCommand
        Dim cmdCr As SqlCommand

        cmdDr = New SqlCommand(strSQLDebit, cn)
        cmdCr = New SqlCommand(strSQLCredit, cn)

        Try
            cn.Open()
            drSQL = cmdDr.ExecuteReader()
            While drSQL.Read()
                GrossTrans(CByte(drSQL.Item(0).ToString.Substring(4, 2)), 1) = FormatNumber(drSQL.Item(1), 2, TriState.True, TriState.False, TriState.False)
                'TotalMoney = TotalMoney + System.Math.Abs(drSQL.Item(1))
            End While
            cn.Close()
            drSQL = Nothing
        Catch ex As Exception
        End Try

        Try
            cn.Open()
            drSQL = cmdCr.ExecuteReader()
            While drSQL.Read()
                GrossTrans(CByte(drSQL.Item(0).ToString.Substring(4, 2)), 2) = FormatNumber(FixNull(drSQL.Item(1)), 2, TriState.True, TriState.False, TriState.False)
                'TotalMoney = TotalMoney + System.Math.Abs(drSQL.Item(1))
            End While
            cn.Close()
            drSQL = Nothing
        Catch ex As Exception
        End Try

        Return GrossTrans

    End Function

    '    Private Function AccountTransactions(ByVal accNo As String, ByVal chkMonth As Months, ByVal DrCr As TranFlows) As String

    '        Dim TaxPeriod As Integer
    '        Dim PeriodStart As String
    '        Dim PeriodEnd As String
    '        Dim strOperator As String

    '        If DrCr = TranFlows.Credit Then
    '            strOperator = ">"
    '        Else
    '            strOperator = "<"
    '        End If

    '        TaxPeriod = CInt(Left(SARSForm.PeriodEnd, 4))
    '        If chkMonth > Months.February Then
    '            TaxPeriod = TaxPeriod - 1
    '        End If
    '        strPeriod = CStr(TaxPeriod) & FixedLengthString(CStr(chkMonth), 2, "Right", "0") & "01"

    '        Dim strSQL As String = "SELECT SUM(tblTransactions.trnAmt) AS GrossTrans " & _
    '                               "FROM   tblAccount INNER JOIN " & _
    '                                      "tblTransactions ON tblAccount.accId = tblTransactions.accId " & _
    '                               "WHERE  tblAccount.accNO = '" & accNo & "' AND " & _
    '                                      "trnTransBatchMonth = '" & strPeriod & "' AND " & _
    '                                      "tblTransactions.trnAmt " & strOperator & "0 AND " & _
    '                                      "tblTransactions.Deleted = 0"

    '        Dim cn As SqlConnection = New SqlConnection(strConnection)
    '        Dim drSQL As SqlDataReader
    '        Dim cmd As SqlCommand

    '        cmd = New SqlCommand(strSQL, cn)
    '        cn.Open()
    '        drSQL = cmd.ExecuteReader()
    '        Try
    '            drSQL.Read()
    '            AccountTransactions = FormatNumber(FixNull(drSQL.Item(0)), 2, TriState.True, TriState.False, TriState.False)
    '        Catch ex As Exception
    '            AccountTransactions = "0.00"
    '            End Try
    '        cn.Close()
    '    End Function

    Private Function FileTrailer(ByVal NoOfRecords As Integer) As String

        Const _134_SEC_ID As String = "T"
        Dim _135_TOTAL_RECORDS As String = NoOfRecords
        Dim _136_MD5 As String = "HASH_TOTAL_" & FileSeq
        Dim _137_TOTAL_MONETARY As String = FormatNumber(TotalMoney, 2, TriState.True, TriState.False, TriState.False)

        ExportFiles(1, FileSeq - 1) = NoOfRecords

        FileTrailer = _134_SEC_ID & DataDelimiter &
                      _135_TOTAL_RECORDS & DataDelimiter &
                      _136_MD5 & DataDelimiter &
                      _137_TOTAL_MONETARY
    End Function

    'Private Function FileTrailer(ByVal NoOfRecords As Integer) As String

    '    FileTrailer = "T" & FixedLengthString(NoOfRecords, 8, "Right", "0")
    'End Function

    'Private Sub WriteIncomeRecords(ByVal cifNO As String, ByVal InterestIncome As Double)
    Private Sub WriteIncomeRecords(ByVal cifNO As String)

        Dim strSQL As String = "SELECT tblTaxCertificate.taxCIF, tblTaxCertificate.taxPaid, tblTaxCertificate.taxAccrued, " &
                                      "tblTaxCertificate.taxAccountNo, tblAccount.atId, tblAccount.accRegDate, " &
                                      "tblAccount.asId, tblAccount.accId " &
                               "FROM   tblAccount INNER JOIN " &
                                      "tblTaxCertificate ON tblAccount.accNO = tblTaxCertificate.taxAccountNo " &
                               "WHERE  tblTaxCertificate.taxCIF = '" & cifNO & "' AND " &
                                      "tblTaxCertificate.taxStartDate = '" & SARSForm.PeriodStart(SARSForm.txtPeriodLength.Text) & "' AND " &
                                      "tblTaxCertificate.taxEndDate = '" & SARSForm.PeriodEnd & "' AND " &
                                      "tblTaxCertificate.taxDeleted = 0 AND " &
                                      "tblAccount.atId = 15"

        'Dim strSQL As String = "SELECT NotImported.cifNO, NotImported.niAmnt, Imported.iAmnt, NotImported.accNO, NotImported.AccType, NotImported.OpenDate, NotImported.Status, NotImported.accId " & _
        '                       "FROM (SELECT tblCIF.cifNO as cifNO, sum(tblAccountMemoInterest.accIntAmt) as niAmnt, tblAccount.accNO as accNO, tblAccount.atId as AccType, tblAccount.accRegDate as OpenDate, tblAccount.asId as Status, tblAccount.accId as accId " & _
        '                             "FROM tblCIF INNER JOIN jblCIFAccount ON tblCIF.cifId = jblCIFAccount.cifId " & _
        '                                         "INNER JOIN tblAccount ON jblCIFAccount.accId = tblAccount.accId " & _
        '                                         "INNER JOIN tblAccountMemoInterest ON tblAccount.accId = tblAccountMemoInterest.accId " & _
        '                             "WHERE (tblCIF.cifId = " & cifId & ") AND (tblAccount.atId < 71) AND " & _
        '                                   "(tblAccountMemoInterest.accIntCapd = 1) AND " & _
        '                                   "(tblAccountMemoInterest.accPostDate >= '" & SARSForm.PeriodStart(SARSForm.txtPeriodLength.Text) & "') AND " & _
        '                                   "(tblAccountMemoInterest.accPostDate <= '" & SARSForm.PeriodEnd & " ') AND " & _
        '                                   "(NOT (tblAccountMemoInterest.accIntRef LIKE '%IMPORT FROM SA THRIFT%')) " & _
        '                             "GROUP BY tblCIF.cifNO,tblAccount.accNO, tblAccount.atId, tblAccount.accRegDate, tblAccount.asId, tblAccount.accId) as NotImported " & _
        '                       "LEFT OUTER JOIN " & _
        '                             "(SELECT tblCIF.cifNO as cifNO, sum(tblAccountMemoInterest.accIntAmt) as iAmnt, tblAccount.accNO as accNO " & _
        '                             "FROM tblCIF INNER JOIN jblCIFAccount ON tblCIF.cifId = jblCIFAccount.cifId " & _
        '                                         "INNER JOIN tblAccount ON jblCIFAccount.accId = tblAccount.accId " & _
        '                                         "INNER JOIN tblAccountMemoInterest ON tblAccount.accId = tblAccountMemoInterest.accId " & _
        '                             "WHERE (tblCIF.cifId = " & cifId & ") AND (tblAccount.atId < 71) AND " & _
        '                                   "(tblAccountMemoInterest.accIntCapd = 1) AND " & _
        '                                   "(tblAccountMemoInterest.accIntRef = 'IMPORT FROM SA THRIFT') " & _
        '                             "GROUP BY tblCIF.cifNO,tblAccount.accNO) AS Imported " & _
        '                       "ON NotImported.accNO = Imported.accNO"

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand

        cmd = New SqlCommand(strSQL, cn)
        cn.Open()
        drSQL = cmd.ExecuteReader()
        While drSQL.Read()
            Dim AccountInfo(drSQL.FieldCount - 1) As Object
            Dim fieldCount As Integer = drSQL.GetValues(AccountInfo)
            WriteToDatabase(IT3sClientDataTable, IncomeDetails(AccountInfo))  'True
            Application.DoEvents()
        End While
        cn.Close()
    End Sub

    Private Function DigitsAdded(ByVal StringOfDigits As String) As Integer

        Dim i As Integer

        For i = 0 To StringOfDigits.Length - 1
            If IsNumeric(StringOfDigits.Substring(i, 1)) Then
                DigitsAdded = DigitsAdded + StringOfDigits.Substring(i, 1)
            End If
        Next
    End Function

    Private Function ClientWithReferences(ByVal CIFInfo As Object) As Boolean

        Dim IT_REF_NO As String
        Dim TP_ID As String
        Dim TP_OTHER_ID As String
        Dim TP_OTHER_OTHER_ID As String
        Dim TP_DOB As String
        Dim CO_REG_NO As String

        ClientWithReferences = False

        Select Case CIFInfo(ClientFields.clnttype)

            Case 1 'Individuals

                IT_REF_NO = FixedLengthString(CIFInfo(ClientFields.cifiTaxNO), 10, "Left", "0", "/", " ")
                TP_ID = FixNull(CIFInfo(ClientFields.cifiIDNO)).Replace(" ", "")
                TP_OTHER_ID = FixNull(CIFInfo(ClientFields.cifiForeignIDNO))
                TP_OTHER_OTHER_ID = FixNull(CIFInfo(ClientFields.cifiPassportNO))
                TP_DOB = FixNull(CIFInfo(ClientFields.cifiDOB))
                If IT_REF_NO & TP_ID & TP_OTHER_ID & TP_OTHER_OTHER_ID & TP_DOB <> "" Then
                    If DigitsAdded(IT_REF_NO & TP_ID & TP_OTHER_ID & TP_OTHER_OTHER_ID & TP_DOB) <> 0 Then
                        ClientWithReferences = True
                    End If
                End If
                If ClientWithReferences Then 'And SARSForm.SkipInvalidReferences Then
                    If IDNoModulusCheck(CIFInfo(ClientFields.cifNO), TP_ID) = "" Then
                        ClientWithReferences = False
                        WriteExceptionToDatabase(IT3sExceptionTable, CIFInfo(ClientFields.cifNO), "References", "Invalid or missing Identity Number")
                    End If
                    If IT_REF_NO = "0000000000" Then
                        WriteExceptionToDatabase(IT3sExceptionTable, CIFInfo(ClientFields.cifNO), "References", "Missing Tax Reference")
                    Else
                        If ValidTaxReference(CIFInfo(ClientFields.cifNO), IT_REF_NO) = "" Then
                            ClientWithReferences = False
                            WriteExceptionToDatabase(IT3sExceptionTable, CIFInfo(ClientFields.cifNO), "References", "Invalid Tax Reference")
                        End If
                    End If
                Else    'If Not ClientWithReferences Then
                    WriteExceptionToDatabase(IT3sExceptionTable, CIFInfo(ClientFields.cifNO), "References", "No usable references")
                End If

            Case 2   'Companies

                IT_REF_NO = FixedLengthString(CIFInfo(ClientFields.cifcTaxNO), 10, "Left", "0", "/", " ")
                CO_REG_NO = FixNull(CIFInfo(ClientFields.cifcRegNO))
                If IT_REF_NO & CO_REG_NO <> "" Then
                    If DigitsAdded(IT_REF_NO & CO_REG_NO) <> 0 Then
                        ClientWithReferences = True
                    End If
                End If
                If ClientWithReferences Then 'And SARSForm.SkipInvalidReferences Then
                    Select Case SARSClientType(CIFInfo(ClientFields.companytype))
                        Case "PRIVATE_CO", "PUBLIC_CO", "OTHER_CO", "CLOSE_CORPORATION"
                            'If ValidTaxReference(CIFInfo(ClientFields.cifNO), IT_REF_NO) = "" And ValidCompanyRegNo(CIFInfo(ClientFields.cifNO), DropChar(DropChar(FixNull(CIFInfo(ClientFields.cifcRegNO)), "/"), " ")) = "" Then
                            If ValidCompanyRegNo(CIFInfo(ClientFields.cifNO), FixNull(CIFInfo(ClientFields.cifcRegNO))) = "" Then
                                ClientWithReferences = False
                                WriteExceptionToDatabase(IT3sExceptionTable, CIFInfo(ClientFields.cifNO), "References", "Invalid or missing Company Registration Number")
                            End If
                        Case "INTERVIVOS_TRUST"
                            'If ValidTaxReference(CIFInfo(ClientFields.cifNO), IT_REF_NO) = "" And ValidTrustRegNo(CIFInfo(ClientFields.cifNO), DropChar(DropChar(FixNull(CIFInfo(ClientFields.cifcRegNO)), "/"), " ")) = "" Then
                            If ValidTrustRegNo(CIFInfo(ClientFields.cifNO), FixNull(CIFInfo(ClientFields.cifcRegNO))) = "" Then
                                ClientWithReferences = False
                                WriteExceptionToDatabase(IT3sExceptionTable, CIFInfo(ClientFields.cifNO), "References", "Invalid or missing Trust Registration Number")
                            End If
                        Case Else
                            If IT_REF_NO = "0000000000" Then
                                WriteExceptionToDatabase(IT3sExceptionTable, CIFInfo(ClientFields.cifNO), "References", "Missing Tax Reference")
                            Else
                                If ValidTaxReference(CIFInfo(ClientFields.cifNO), IT_REF_NO) = "" Then
                                    ClientWithReferences = False
                                    WriteExceptionToDatabase(IT3sExceptionTable, CIFInfo(ClientFields.cifNO), "References", "Invalid Tax Reference")
                                End If
                            End If
                    End Select
                ElseIf Not ClientWithReferences Then
                    WriteExceptionToDatabase(IT3sExceptionTable, CIFInfo(ClientFields.cifNO), "References", "No usable references")
                End If
        End Select
    End Function

    'Private Function ClientWithReferences(ByVal CIFInfo As Object) As Boolean

    '    Dim IT_REF_NO As String
    '    Dim TP_ID As String
    '    Dim TP_OTHER_ID As String
    '    Dim TP_OTHER_OTHER_ID As String
    '    Dim TP_DOB As String
    '    Dim CO_REG_NO As String

    '    ClientWithReferences = False
    '    Select Case CIFInfo(ClientFields.clnttype)
    '        Case 1 'Individuals
    '            IT_REF_NO = FixedLengthString(CIFInfo(ClientFields.cifiTaxNO), 10, "Left", "0", "/", " ")
    '            TP_ID = FixNull(CIFInfo(ClientFields.cifiIDNO)).Replace(" ", "")
    '            TP_OTHER_ID = FixNull(CIFInfo(ClientFields.cifiForeignIDNO))
    '            TP_OTHER_OTHER_ID = FixNull(CIFInfo(ClientFields.cifiPassportNO))
    '            TP_DOB = FixNull(CIFInfo(ClientFields.cifiDOB))
    '            If IT_REF_NO & TP_ID & TP_OTHER_ID & TP_OTHER_OTHER_ID & TP_DOB <> "" Then
    '                If DigitsAdded(IT_REF_NO & TP_ID & TP_OTHER_ID & TP_OTHER_OTHER_ID & TP_DOB) <> 0 Then
    '                    ClientWithReferences = True
    '                End If
    '            End If
    '            If ClientWithReferences And SARSForm.SkipInvalidReferences Then
    '                If ValidTaxReference(CIFInfo(ClientFields.cifNO), IT_REF_NO) = "" And IDNoModulusCheck(CIFInfo(ClientFields.cifNO), TP_ID) = "" Then
    '                    ClientWithReferences = False
    '                    WriteExceptionToDatabase(IT3sExceptionTable, CIFInfo(ClientFields.cifNO), "References", "Client skipped - invalid or missing Tax Reference or Identity Number")
    '                End If
    '            ElseIf Not ClientWithReferences Then
    '                WriteExceptionToDatabase(IT3sExceptionTable, CIFInfo(ClientFields.cifNO), "References", "Client skipped - no usable references")
    '            End If
    '        Case 2   'Companies
    '            IT_REF_NO = FixedLengthString(CIFInfo(ClientFields.cifcTaxNO), 10, "Left", "0", "/", " ")
    '            CO_REG_NO = FixNull(CIFInfo(ClientFields.cifcRegNO))
    '            If IT_REF_NO & CO_REG_NO <> "" Then
    '                If DigitsAdded(IT_REF_NO & CO_REG_NO) <> 0 Then
    '                    ClientWithReferences = True
    '                End If
    '            End If
    '            If ClientWithReferences And SARSForm.SkipInvalidReferences Then
    '                Select Case SARSClientType(CIFInfo(ClientFields.companytype))
    '                    Case "PRIVATE_CO", "PUBLIC_CO", "OTHER_CO"
    '                        'If ValidTaxReference(CIFInfo(ClientFields.cifNO), IT_REF_NO) = "" And ValidCompanyRegNo(CIFInfo(ClientFields.cifNO), DropChar(DropChar(FixNull(CIFInfo(ClientFields.cifcRegNO)), "/"), " ")) = "" Then
    '                        If ValidCompanyRegNo(CIFInfo(ClientFields.cifNO), DropChar(DropChar(FixNull(CIFInfo(ClientFields.cifcRegNO)), "/"), " ")) = "" Then
    '                            ClientWithReferences = False
    '                            WriteExceptionToDatabase(IT3sExceptionTable, CIFInfo(ClientFields.cifNO), "References", "Client skipped - invalid or missing Company Registration Number")
    '                        End If
    '                    Case "INTERVIVOS_TRUST"
    '                        'If ValidTaxReference(CIFInfo(ClientFields.cifNO), IT_REF_NO) = "" And ValidTrustRegNo(CIFInfo(ClientFields.cifNO), DropChar(DropChar(FixNull(CIFInfo(ClientFields.cifcRegNO)), "/"), " ")) = "" Then
    '                        If ValidTrustRegNo(CIFInfo(ClientFields.cifNO), DropChar(DropChar(FixNull(CIFInfo(ClientFields.cifcRegNO)), "/"), " ")) = "" Then
    '                            ClientWithReferences = False
    '                            WriteExceptionToDatabase(IT3sExceptionTable, CIFInfo(ClientFields.cifNO), "References", "Client skipped - invalid or missing Trust Registration Number")
    '                        End If
    '                    Case Else
    '                        If ValidTaxReference(CIFInfo(ClientFields.cifNO), IT_REF_NO) = "" Then
    '                            ClientWithReferences = False
    '                            WriteExceptionToDatabase(IT3sExceptionTable, CIFInfo(ClientFields.cifNO), "References", "Client skipped - invalid or missing Tax Reference or Company Registration Number")
    '                        Else
    '                            ClientWithReferences = True
    '                        End If
    '                End Select
    '            ElseIf Not ClientWithReferences Then
    '                WriteExceptionToDatabase(IT3sExceptionTable, CIFInfo(ClientFields.cifNO), "References", "Client skipped - no usable references")
    '            End If
    '    End Select
    'End Function

    Private Function GetNoOfClientAccounts(ByVal cifNO As String) As Integer

        ''"Normal" selection
        'Dim strSQL As String = "SELECT COUNT(taxAccountNo) AS NoOfAccounts " & _
        '                       "FROM   tblTaxCertificate " & _
        '                       "WHERE  taxCIF = '" & cifNO & "' AND " & _
        '                              "taxStartDate = '" & SARSForm.PeriodStart(SARSForm.txtPeriodLength.Text) & "' AND " & _
        '                              "taxEndDate = '" & SARSForm.PeriodEnd & "' AND " & _
        '                              "taxDeleted = 0"

        'Selection for Tax-free Savings Accounts, including the Corporate Savers masquerading as Tax-free Savings Accounts
        'during the period between introduction of the product and switchover to the real thing in early 2016
        '21/10/2016:  Actually we have to exclude the Corporate Savers as the transfer transactions are causing rejections on the submission
        Dim strSQL As String = "SELECT COUNT(tblTaxCertificate.taxAccountNo) AS NoOfAccounts " &
                               "FROM   tblTaxCertificate INNER JOIN " &
                                      "tblAccount ON tblTaxCertificate.taxAccountNo = tblAccount.accNO " &
                               "WHERE  taxCIF = '" & cifNO & "' AND " &
                                      "taxStartDate = '" & SARSForm.PeriodStart(SARSForm.txtPeriodLength.Text) & "' AND " &
                                      "taxEndDate = '" & SARSForm.PeriodEnd & "' AND " &
                                      "taxDeleted = 0 AND " &
                                      "tblAccount.atId = 15"

        'Dim strSQL As String = "SELECT * FROM " & _
        '                        "(SELECT tblAccount.accNO " & _
        '                         "FROM jblCIFAccount INNER JOIN tblAccount ON jblCIFAccount.accId = tblAccount.accId " & _
        '                                            "INNER JOIN tblAccountMemoInterest ON tblAccount.accId = tblAccountMemoInterest.accId " & _
        '                         "WHERE(jblCIFAccount.cifId = " & cifId & ") And " & _
        '                              "(tblAccount.atId < 71) And " & _
        '                              "(tblAccountMemoInterest.accIntCapd = 1) And " & _
        '                              "(tblAccountMemoInterest.accPostDate >= '" & SARSForm.PeriodStart(SARSForm.txtPeriodLength.Text) & "') AND " & _
        '                              "(tblAccountMemoInterest.accPostDate <= '" & SARSForm.PeriodEnd & " ') AND " & _
        '                              "(NOT (tblAccountMemoInterest.accIntRef LIKE '%IMPORT FROM SA THRIFT%')) AND " & _
        '                              "(tblAccountMemoInterest.accIntAmt <> 0) " & _
        '                         "GROUP BY tblAccount.accNO) AS NotImported " & _
        '                       "LEFT OUTER JOIN " & _
        '                        "(SELECT tblAccount.accNO " & _
        '                         "FROM jblCIFAccount INNER JOIN tblAccount ON jblCIFAccount.accId = tblAccount.accId " & _
        '                                            "INNER JOIN tblAccountMemoInterest ON tblAccount.accId = tblAccountMemoInterest.accId " & _
        '                         "WHERE (jblCIFAccount.cifId = " & cifId & ") AND " & _
        '                               "(tblAccount.atId < 71) And " & _
        '                               "(tblAccountMemoInterest.accIntCapd = 1) AND " & _
        '                               "(tblAccountMemoInterest.accIntRef = 'IMPORT FROM SA THRIFT') AND " & _
        '                               "(tblAccountMemoInterest.accIntAmt <> 0) " & _
        '                         "GROUP BY tblAccount.accNO) AS Imported " & _
        '                       "ON Imported.accNO = NotImported.accNO"

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand

        cmd = New SqlCommand(strSQL, cn)
        GetNoOfClientAccounts = 0
        cn.Open()
        drSQL = cmd.ExecuteReader()
        If drSQL.Read Then
            GetNoOfClientAccounts = drSQL.Item(0)
        End If
        cn.Close()
    End Function

    'Private Function GetInterestIncome(ByVal accNO As String) As Double

    '    Dim tmpInterestIncome As Double = 0

    '    Dim strSQL As String = "SELECT taxPaid + taxAccrued AS TotalTax " & _
    '                           "FROM   tblTaxCertificate " & _
    '                           "WHERE  taxAccountNo = '" & accNO & "' AND " & _
    '                                  "taxStartDate = '" & SARSForm.PeriodStart(SARSForm.txtPeriodLength.Text) & "' AND " & _
    '                                  "taxEndDate = '" & SARSForm.PeriodEnd & "' AND " & _
    '                                  "taxDeleted = 0"

    '    'Dim strSQL As String = "SELECT Imported.iAmnt + NotImported.niAmnt " & _
    '    '                       "FROM " & _
    '    '                        "(SELECT jblCIFAccount.cifId, SUM(tblAccountMemoInterest.accIntAmt) AS iAmnt " & _
    '    '                         "FROM jblCIFAccount INNER JOIN tblAccount ON jblCIFAccount.accId = tblAccount.accId " & _
    '    '                                            "INNER JOIN tblAccountMemoInterest ON tblAccount.accId = tblAccountMemoInterest.accId " & _
    '    '                         "WHERE (jblCIFAccount.cifId = " & cifId & ") AND " & _
    '    '                               "(tblAccount.atId < 71) AND " & _
    '    '                               "(tblAccountMemoInterest.accIntCapd = 1) AND " & _
    '    '                               "(tblAccountMemoInterest.accIntRef = 'IMPORT FROM SA THRIFT') AND " & _
    '    '                               "(tblAccountMemoInterest.accIntAmt <> 0) " & _
    '    '                         "GROUP BY jblCIFAccount.cifId ) AS Imported " & _
    '    '                       "INNER JOIN " & _
    '    '                        "(SELECT jblCIFAccount.cifId, SUM(tblAccountMemoInterest.accIntAmt) AS niAmnt " & _
    '    '                         "FROM jblCIFAccount INNER JOIN tblAccount ON jblCIFAccount.accId = tblAccount.accId " & _
    '    '                                            "INNER JOIN tblAccountMemoInterest ON tblAccount.accId = tblAccountMemoInterest.accId " & _
    '    '                         "WHERE (jblCIFAccount.cifId = " & cifId & ") AND " & _
    '    '                               "(tblAccount.atId < 71) AND " & _
    '    '                               "(tblAccountMemoInterest.accIntCapd = 1) AND " & _
    '    '                               "(tblAccountMemoInterest.accPostDate >= '" & SARSForm.PeriodStart(SARSForm.txtPeriodLength.Text) & "') AND " & _
    '    '                               "(tblAccountMemoInterest.accPostDate <= '" & SARSForm.PeriodEnd & " ') AND " & _
    '    '                               "(NOT (tblAccountMemoInterest.accIntRef LIKE '%IMPORT FROM SA THRIFT%')) AND " & _
    '    '                               "(tblAccountMemoInterest.accIntAmt <> 0) " & _
    '    '                         "GROUP BY jblCIFAccount.cifId ) AS NotImported " & _
    '    '                       "ON Imported.cifId = NotImported.cifId"

    '    Dim cn As SqlConnection = New SqlConnection(strConnection)
    '    Dim drSQL As SqlDataReader
    '    Dim cmd As SqlCommand

    '    cmd = New SqlCommand(strSQL, cn)
    '    cn.Open()
    '    drSQL = cmd.ExecuteReader()
    '    If drSQL.Read() Then
    '        tmpInterestIncome = drSQL.Item(0)
    '    End If
    '    cn.Close()
    '    TotalMoney = TotalMoney + tmpInterestIncome
    '    Return tmpInterestIncome
    'End Function

    Public Sub CreateExportData()

        Dim strSQL As String = "SELECT	*
                                FROM	tblCIF LEFT OUTER JOIN
		                                tblCIFIndividual ON tblCIF.cifId = tblCIFIndividual.cifId LEFT OUTER JOIN
		                                tblCIFCompany ON tblCIF.cifId = tblCIFCompany.cifId
                                WHERE	tblCIF.cifId IN	   (SELECT tblCIF.cifId
							                                FROM   tblCIF LEFT OUTER JOIN 
									                                tblCIFIndividual ON tblCIF.cifId = tblCIFIndividual.cifId LEFT OUTER JOIN 
									                                tblCIFCompany ON tblCIF.cifId = tblCIFCompany.cifId INNER JOIN 
									                                jblCIFAccount ON tblCIF.cifId = jblCIFAccount.cifId
                                                                                                        AND jblCIFAccount.Deleted = 0 INNER JOIN 
									                                tblAccount ON jblCIFAccount.accId = tblAccount.accId 
							                                WHERE  tblAccount.atId = 15
										                                AND tblCIF.Deleted = 0)"
        'And tblCIF.cifNO = '031693')"

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand
        Dim NoOfClientAccounts As Integer
        Dim PersonalDetails As String
        Dim strUniqueID As String
        Dim ClientHasReferences As Boolean
        'Dim InterestIncome As Double

        SARSForm.SetButton = "Stop"
        Exporting = True
        Abort = False
        TotalRecords = 0
        Skipped = 0
        ClientsSkippedDueToPersonalInfo = 0
        ClientsSkippedDueToReferences = 0
        ClientsSkippedDueInactivity = 0
        ClientsSkippedDueToList = 0
        FileSeq = 1
        RecSeq = 0
        TotalMoney = 0
        SARSFile = BuildFileName()
        'WriteToDatabase(IT3sHeaderTable, FileHeader(SARSFile.Substring(SARSFile.LastIndexOf("\") + 1).Substring(0, SARSFile.Substring(SARSFile.LastIndexOf("\") + 1).Length - 4).Replace(FileNameDelimiter, "-")))
        strUniqueID = SARSFile.Substring(0, SARSFile.LastIndexOf(FileNameDelimiter))
        strUniqueID = strUniqueID.Substring(strUniqueID.LastIndexOf(FileNameDelimiter) + 1)
        WriteToDatabase(IT3sHeaderTable, FileHeader(strUniqueID))
        WriteToDatabase(IT3sHeaderTable, SubmittingEntityDetails, False)
        cmd = New SqlCommand(strSQL, cn)
        cn.Open()
        drSQL = cmd.ExecuteReader()
        While drSQL.Read() And Not Abort
            InvalidIDExceptionRaised = False
            InvalidTaxNoExceptionRaised = False
            InvalidCoRegNoExceptionRaised = False
            InvalidTrustRegNoExceptionRaised = False
            Dim CIFInfo(drSQL.FieldCount - 1) As Object
            Dim fieldCount As Integer = drSQL.GetValues(CIFInfo)
            If Not (CIFToSkip(CIFInfo(ClientFields.cifNO))) Then
                NoOfClientAccounts = GetNoOfClientAccounts(CIFInfo(ClientFields.cifNO))
                If NoOfClientAccounts > 0 Then
                    ClientHasReferences = ClientWithReferences(CIFInfo)
                    If ClientHasReferences Or (Not ClientHasReferences And Not SARSForm.SkipInvalidReferences) Then
                        PersonalDetails = PersonalDetailsRecord(CIFInfo, ClientHasReferences)
                        If PersonalDetails <> "" Then
                            'InterestIncome = GetInterestIncome(CIFInfo(ClientFields.cifNO))
                            If RecSeq > MaxFileRec - NoOfClientAccounts - 4 Then        '4 = Header + Entity Data + Trailer + Personal Info Record
                                WriteToDatabase(IT3sHeaderTable, FileTrailer(RecSeq - 2), False)
                                FileSeq = FileSeq + 1
                                RecSeq = 0
                                TotalMoney = 0
                                SARSFile = BuildFileName()
                                'WriteToDatabase(IT3sHeaderTable, FileHeader(SARSFile.Substring(SARSFile.LastIndexOf("\") + 1).Substring(0, SARSFile.Substring(SARSFile.LastIndexOf("\") + 1).Length - 4).Replace("_", "-")), False)
                                strUniqueID = SARSFile.Substring(0, SARSFile.LastIndexOf(FileNameDelimiter))
                                strUniqueID = strUniqueID.Substring(strUniqueID.LastIndexOf(FileNameDelimiter) + 1)
                                WriteToDatabase(IT3sHeaderTable, FileHeader(strUniqueID), False)
                                WriteToDatabase(IT3sHeaderTable, SubmittingEntityDetails)
                            End If
                            WriteToDatabase(IT3sClientDataTable, PersonalDetails)  'true
                            Application.DoEvents()
                            WriteIncomeRecords(CIFInfo(ClientFields.cifNO))
                            SARSForm.ClientRecords = TotalRecords
                        Else
                            Skipped = Skipped + 1
                            ClientsSkippedDueToPersonalInfo = ClientsSkippedDueToPersonalInfo + 1
                            SARSForm.ClientsSkipped = Skipped
                            SARSForm.ClientsSkippedDueToPersonalInfo = ClientsSkippedDueToPersonalInfo
                        End If
                    Else
                        Skipped = Skipped + 1
                        ClientsSkippedDueToReferences = ClientsSkippedDueToReferences + 1
                        SARSForm.ClientsSkipped = Skipped
                        SARSForm.ClientsSkippedDueToReferences = ClientsSkippedDueToReferences
                    End If
                Else
                    Skipped = Skipped + 1
                    ClientsSkippedDueInactivity = ClientsSkippedDueInactivity + 1
                    SARSForm.ClientsSkipped = Skipped
                    SARSForm.ClientsSkippedDueInactivity = ClientsSkippedDueInactivity
                    WriteExceptionToDatabase(IT3sExceptionTable, CIFInfo(ClientFields.cifNO), "Skip", "Client skipped - no active accounts")
                End If
            Else
                Skipped = Skipped + 1
                ClientsSkippedDueToList = ClientsSkippedDueToList + 1
                SARSForm.ClientsSkipped = Skipped
                SARSForm.ClientsSkippedDueToList = ClientsSkippedDueToList
                WriteExceptionToDatabase(IT3sExceptionTable, CIFInfo(ClientFields.cifNO), "Skip", "Client flagged to be skipped")
            End If
            Application.DoEvents()
        End While
        WriteToDatabase(IT3sHeaderTable, FileTrailer(RecSeq - 2), False)
        cn.Close()
        Exporting = False
        SARSForm.SetButton = "Exit"
    End Sub

    'Public Sub CreateExportFile()

    '    Dim strSQL As String = "SELECT * FROM tblCIF LEFT OUTER JOIN tblCIFIndividual ON tblCIF.cifId = tblCIFIndividual.cifId " & _
    '                           "LEFT OUTER JOIN tblCIFCompany ON tblCIF.cifId = tblCIFCompany.cifId"
    '    'Dim strSQL As String = "SELECT * FROM tblCIF LEFT OUTER JOIN tblCIFIndividual ON tblCIF.cifId = tblCIFIndividual.cifId " & _
    '    '                       "LEFT OUTER JOIN tblCIFCompany ON tblCIF.cifId = tblCIFCompany.cifId " & _
    '    '                       "WHERE tblCIF.cifNO = '040494'" ' & _
    '    '"tblCIF.cifNO = '037389' OR " & _
    '    '"tblCIF.cifNO = '038122' OR " & _
    '    '"tblCIF.cifNO = '034691' OR " & _
    '    '"tblCIF.cifNO = '029848' OR " & _
    '    '"tblCIF.cifNO = '020921'"
    '    Dim cn As SqlConnection = New SqlConnection(strConnection)
    '    Dim drSQL As SqlDataReader
    '    Dim cmd As SqlCommand
    '    Dim NoOfClientAccounts As Integer
    '    Dim InterestIncome As Double

    '    StartTime = Now
    '    SARSForm.SetButton = "Stop"
    '    Exporting = True
    '    Abort = False
    '    TotalRecords = 0
    '    Skipped = 0
    '    FileSeq = 1
    '    RecSeq = 1
    '    SARSFile = BuildFileName()
    '    If File.Exists(SARSFile) Then
    '        File.Delete(SARSFile)
    '    End If
    '    WriteToDatabase(IT3sHeaderTable, FileHeader(SARSFile.Substring(SARSFile.LastIndexOf("\") + 1)), False)
    '    WriteToDatabase(IT3sHeaderTable, SubmittingEntityDetails)
    '    cmd = New SqlCommand(strSQL, cn)
    '    cn.Open()
    '    drSQL = cmd.ExecuteReader()
    '    While drSQL.Read() And Not Abort
    '        Dim CIFInfo(drSQL.FieldCount - 1) As Object
    '        Dim fieldCount As Integer = drSQL.GetValues(CIFInfo)
    '        If Not (CIFToSkip(CIFInfo(ClientFields.cifNO))) And ClientWithReferences(CIFInfo) Then
    '            NoOfClientAccounts = GetNoOfClientAccounts(drSQL.Item(4))
    '            InterestIncome = GetInterestIncome(drSQL.Item(4))
    '            If NoOfClientAccounts > 0 And InterestIncome > 0 Then
    '                If RecSeq > MaxFileRec - NoOfClientAccounts - 3 Then        '3 = Header + Trailer + Personal Info Record
    '                    WriteToDatabase(IT3sClientDataTable, FileTrailer(RecSeq - 1), False)
    '                    FileSeq = FileSeq + 1
    '                    RecSeq = 0
    '                    SARSFile = BuildFileName()
    '                    If File.Exists(SARSFile) Then
    '                        File.Delete(SARSFile)
    '                    End If
    '                    WriteToDatabase(IT3sHeaderTable, FileHeader(SARSFile.Substring(SARSFile.LastIndexOf("\") + 1)), False)
    '                End If
    '                WriteToDatabase(IT3sClientDataTable, PersonalDetailsRecord(CIFInfo))  'true
    '                WriteIncomeRecords(drSQL.Item(4))
    '                SARSForm.ClientRecords = TotalRecords
    '            Else
    '                Skipped = Skipped + 1
    '                SARSForm.ClientsSkipped = Skipped
    '            End If
    '        Else
    '            Skipped = Skipped + 1
    '            SARSForm.ClientsSkipped = Skipped
    '        End If
    '        Application.DoEvents()
    '    End While
    '    WriteToDatabase(IT3sClientDataTable, FileTrailer(RecSeq - 1), False)
    '    cn.Close()
    '    Exporting = False
    '    SARSForm.SetButton = "Exit"
    'End Sub

    Public Sub CreateExportFiles(SubmissionPeriod As String, SubmissionNo As Integer, TestFlag As String)

        Dim i As Byte
        Dim Unique_No As Integer = 0
        Dim AHFD_Unique_No As String
        Dim ExportFile As StreamWriter
        Dim strSQL As String
        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand

        Dim recHeader As String
        Dim recEntity As String
        Dim recTrailer(1) As String

        Dim strResponse As String
        Dim strID As String
        Dim strWrite As String

        'GetMD5HashByLine("\\DomainCtrl\GBSDocs$\DJBSmith\SARS IT3(b) Interface\2013\Test\I3B_1_9110002715_GBS20130228038022001_20130620T173829.txt")
        'GetMD5HashByChar("\\DomainCtrl\GBSDocs$\DJBSmith\SARS IT3(b) Interface\2013\Test\I3B_1_9110002715_GBS20130228038022001_20130620T173829.txt")
        'GetMD5Hash("\\DomainCtrl\GBSDocs$\DJBSmith\SARS IT3(b) Interface\2013\Test\I3B_1_9110002715_GBS20130228017030001_20130620T111733a.txt")

        Unique_No = GetTotalPreviousRecords(it3sPeriod, it3sSubmissionNo - 1, SARSForm.TestFile)

        For i = 1 To FileSeq

            Rec_No = 0

            ExportFile = New StreamWriter(ExportFiles(0, i - 1))

            strSQL = "SELECT hdrId, it3sRecord  " &
                     "FROM " & IT3sHeaderTable & " " &
                     "WHERE  it3sFileNo = " & i & " " &
                                "AND it3sPeriod = '" & it3sPeriod & "' " &
                                "AND it3sSubmissionNo = " & it3sSubmissionNo & " " &
                                "AND TestRun = '" & TestFlag & "'"

            cmd = New SqlCommand(strSQL, cn)
            Try
                cn.Open()
                drSQL = cmd.ExecuteReader()
                While drSQL.Read()
                    Select Case drSQL.Item(1).ToString.Substring(0, 1)
                        Case "H"
                            Select Case drSQL.Item(1).ToString.Substring(2, 2)
                                Case "GH"
                                    recHeader = drSQL.Item(1).ToString.Replace("GROUP_TOTAL", FileSeq)
                                Case "SE"
                                    recEntity = drSQL.Item(1).ToString
                            End Select
                        Case "T"
                            recTrailer(0) = drSQL.Item(0).ToString
                            recTrailer(1) = drSQL.Item(1).ToString
                    End Select
                End While
            Catch ex As Exception
                MessageBox.Show("Header and Trailer information for File No " & i & "incomplete", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            cn.Close()
            cmd = Nothing

            ExportFile.WriteLine(recHeader)
            ExportFile.WriteLine(recEntity)

            strSQL = "SELECT recId, it3sRecord  " &
                     "FROM " & IT3sClientDataTable & " " &
                     "WHERE  it3sFileNo = " & i & " AND " &
                            "it3sRecord LIKE '%|AHDD|%' " &
                                "AND it3sPeriod = '" & it3sPeriod & "' " &
                                "AND it3sSubmissionNo = " & it3sSubmissionNo & " " &
                                "AND TestRun = '" & TestFlag & "' " &
                     "ORDER BY recId ASC"

            cmd = New SqlCommand(strSQL, cn)
            Try
                cn.Open()
                drSQL = cmd.ExecuteReader()
                While drSQL.Read()
                    If InStr(drSQL.Item(1).ToString, "Unique_No") > 0 Then
                        Rec_No += 1
                        Unique_No += 1
                        ExportFile.WriteLine(drSQL.Item(1).ToString.Replace("Unique_No", Unique_No).Replace("Rec_No", Rec_No))
                        RecordUniqueNumbers(drSQL.Item(0), Unique_No)
                    Else
                        strResponse = GetResponse(it3sPeriod, it3sSubmissionNo, TestFlag, GetColumnFromRecord(4, drSQL.Item(1).ToString, DataDelimiter))
                        If InStr(strResponse, DataDelimiter & "R" & DataDelimiter) = 0 Then
                            Rec_No += 1
                            ExportFile.WriteLine(drSQL.Item(1).ToString.Replace("Rec_No", Rec_No))
                        End If
                    End If
                End While
            Catch ex As Exception
                MessageBox.Show("Error encountered while processing personal information records for File No " & i, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            cn.Close()
            cmd = Nothing

            strSQL = "SELECT recId, it3sRecord  " &
                     "FROM " & IT3sClientDataTable & " " &
                     "WHERE  it3sFileNo = " & i & " AND " &
                            "it3sRecord LIKE '%|AHFD|%' " &
                                "AND it3sPeriod = '" & it3sPeriod & "' " &
                                "AND it3sSubmissionNo = " & it3sSubmissionNo & " " &
                                "AND TestRun = '" & TestFlag & "' " &
                     "ORDER BY recId ASC"

            cmd = New SqlCommand(strSQL, cn)
            Try
                cn.Open()
                drSQL = cmd.ExecuteReader()
                While drSQL.Read()
                    If InStr(drSQL.Item(1).ToString, "Unique_No") > 0 Then
                        Rec_No += 1
                        Unique_No += 1
                        ExportFile.WriteLine(drSQL.Item(1).ToString.Replace("Unique_No", Unique_No).Replace("Rec_No", Rec_No))
                        RecordUniqueNumbers(drSQL.Item(0), Unique_No)
                        'InsertUniqueNumbers(drSQL.Item(0), Unique_No, drSQL.Item(1).ToString)
                    Else
                        strResponse = GetResponse(it3sPeriod, it3sSubmissionNo, TestFlag, GetColumnFromRecord(4, drSQL.Item(1).ToString, DataDelimiter))
                        If InStr(strResponse, DataDelimiter & "R" & DataDelimiter) = 0 Then
                            Rec_No += 1
                            ExportFile.WriteLine(drSQL.Item(1).ToString.Replace("Rec_No", Rec_No))
                        End If
                    End If
                End While
            Catch ex As Exception
                MessageBox.Show("Error encountered while processing financial detail records for File No " & i, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            cn.Close()
            cmd = Nothing

            strSQL = "SELECT recId, it3sRecord  " &
                     "FROM " & IT3sClientDataTable & " " &
                     "WHERE  it3sFileNo = " & i & " AND " &
                            "it3sRecord LIKE '%|ATD|%' " &
                                "AND it3sPeriod = '" & it3sPeriod & "' " &
                                "AND it3sSubmissionNo = " & it3sSubmissionNo & " " &
                                "AND TestRun = '" & TestFlag & "' " &
                     "ORDER BY recId ASC"

            cmd = New SqlCommand(strSQL, cn)
            Try
                cn.Open()
                drSQL = cmd.ExecuteReader()
                While drSQL.Read()
                    If InStr(drSQL.Item(1).ToString, "|Unique_No|") > 0 Then
                        Rec_No += 1
                        Unique_No += 1
                        strID = GetColumnFromRecord(7, drSQL.Item(1).ToString, DataDelimiter)
                        AHFD_Unique_No = GetAHFDUniqueNo(strID, TestFlag)
                        strWrite = drSQL.Item(1).ToString.Replace("|Unique_No|", "|" & Unique_No & "|").Replace("Rec_No", Rec_No).Replace("AHFD_Unique_No", AHFD_Unique_No)
                        If strWrite.Contains("<") Then
                            strWrite = strWrite.Replace(strID, "")
                        End If
                        ExportFile.WriteLine(strWrite)
                        RecordUniqueNumbers(drSQL.Item(0), Unique_No)
                        RecordAHFDUniqueNumber(drSQL.Item(0), AHFD_Unique_No)

                    Else
                        strResponse = GetResponse(it3sPeriod, it3sSubmissionNo, TestFlag, GetColumnFromRecord(4, drSQL.Item(1).ToString, DataDelimiter))
                        If InStr(strResponse, DataDelimiter & "R" & DataDelimiter) = 0 Then
                            Rec_No += 1
                            ExportFile.WriteLine(drSQL.Item(1).ToString.Replace("Rec_No", Rec_No))
                        End If
                    End If
                End While
            Catch ex As Exception
                MessageBox.Show("Error encountered while processing transaction detail records for File No " & i, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            cn.Close()
            cmd = Nothing

            ExportFile.Close()
            ExportFile = Nothing

            UpdateUniqueNumbers()

            WriteTrailerWithHashTotal(i, ExportFiles(0, i - 1), recTrailer(0), recTrailer(1))

        Next

    End Sub

    Private Function GetTotalPreviousRecords(SubmissionPeriod As String, SubmissionNo As Integer, TestFlag As String) As Integer

        Dim it3sRecord As String = ""

        Dim strSQL As String = "SELECT it3sRecord " &
                               "FROM   " & IT3sHeaderTable & " " &
                               "WHERE  it3sPeriod = '" & SubmissionPeriod & "' " &
                                            "AND it3sSubmissionNo = " & SubmissionNo & " " &
                                            "AND TestRun = '" & TestFlag & "' " &
                                            "AND LEFT(it3sRecord, 2) = 'T" & DataDelimiter & "'"

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)

        cn.Open()
        drSQL = cmd.ExecuteReader()
        Try
            drSQL.Read()
            it3sRecord = drSQL.Item(0)
        Catch ex As Exception
        End Try
        cn.Close()

        If it3sRecord <> "" Then
            Return CInt(GetColumnFromRecord(2, it3sRecord, DataDelimiter))
        Else
            Return 0
        End If

    End Function

    Private Function Getit3sRecordText(recId As Integer) As String

        Dim it3sRecord As String = ""

        Dim strSQL As String = "SELECT it3sRecord " &
                               "FROM   " & IT3sClientDataTable & " " &
                               "WHERE  recId = " & recId

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)

        cn.Open()
        drSQL = cmd.ExecuteReader()
        Try
            drSQL.Read()
            it3sRecord = drSQL.Item(0)
        Catch ex As Exception
        End Try
        cn.Close()

        Return it3sRecord

    End Function

    Private Function GetUniqueNumber(recId As Integer) As String

        Dim strSQL As String = "SELECT it3sUniqueNo
                                FROM   " & IT3sUniqueNoTable & "
                                WHERE   it3srecId = " & recId.ToString

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)
        Dim UniqueNo As Integer = 0

        Try
            cn.Open()
            drSQL = cmd.ExecuteReader()
            If drSQL.Read() Then
                UniqueNo = drSQL.Item(0)
            End If
        Catch ex As Exception
        End Try
        cn.Close()

        Return UniqueNo.ToString

    End Function

    Private Sub UpdateUniqueNumbers()

        Dim strSQL As String = "SELECT it3srecId, it3sUniqueNo " &
                               "FROM   " & IT3sUniqueNoTable

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)

        Try
            cn.Open()
            drSQL = cmd.ExecuteReader()
            While drSQL.Read()
                InsertUniqueNumbers(drSQL.Item(0), drSQL.Item(1), Getit3sRecordText(drSQL.Item(0)))
            End While
        Catch ex As Exception
        End Try
        cn.Close()

        ClearUniqueNumbers()

    End Sub

    Private Sub RecordUniqueNumbers(recId As Integer, Unique_No As Integer)

        Dim cn As New SqlConnection(strConnection)
        Dim strSQL As String
        Dim cmd As SqlCommand

        strSQL = "INSERT INTO " & IT3sUniqueNoTable & " (it3srecId, it3sUniqueNo) " &
                                                "VALUES (" & recId & ", " & Unique_No & ")"

        cmd = New SqlCommand(strSQL, cn)
        Try
            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()
        Catch ex As Exception
        End Try

    End Sub

    Private Sub RecordAHFDUniqueNumber(recId As Integer, AHFD_Unique_No As String)

        Dim cn As New SqlConnection(strConnection)
        Dim strSQL As String
        Dim cmd As SqlCommand

        strSQL = "UPDATE " & IT3sClientDataTable & "
                  SET   it3sRecord = REPLACE(it3sRecord, '|AHFD_Unique_No|', '|" & AHFD_Unique_No & "|')
                  WHERE recId = " & recId.ToString

        cmd = New SqlCommand(strSQL, cn)
        Try
            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()
        Catch ex As Exception
        End Try

    End Sub

    Private Sub InsertUniqueNumbers(recId As Integer, Unique_No As Integer, it3sRecord As String)

        Dim cn As New SqlConnection(strConnection)
        Dim strSQL As String
        Dim cmd As SqlCommand

        strSQL = "UPDATE " & IT3sClientDataTable & " " &
                 "SET it3sRecord = '" & it3sRecord.Replace("|Unique_No|", "|" & Unique_No & "|") & "' " &
                 "WHERE recId = " & recId
        cmd = New SqlCommand(strSQL, cn)
        Try
            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()
        Catch ex As Exception
            MsgBox("Error in InsertUniqueNumbers", MsgBoxStyle.OkOnly, "Error")
        End Try

    End Sub

    'Private Sub InsertUniqueNumbers(recId As Integer, Unique_No As Integer, it3sRecord As String)

    '    Dim cn As New SqlConnection(strConnection)
    '    Dim strSQL As String
    '    Dim cmd As SqlCommand

    '    strSQL = "UPDATE " & IT3sClientDataTable & " " & _
    '             "SET it3sRecord = N'" & it3sRecord.Replace("Unique_No", Unique_No) & "' " & _
    '             "WHERE recId = " & recId
    '    cmd = New SqlCommand(strSQL, cn)
    '    Try
    '        cn.Open()
    '        cmd.ExecuteNonQuery()
    '        cn.Close()
    '    Catch ex As Exception
    '        MsgBox("Error in InsertUniqueNumbers", MsgBoxStyle.OkOnly, "Error")
    '    End Try

    'End Sub

    Public Sub InsertResponse(it3sSubmissionPeriod As String, it3sSubmissionNo As Integer, TestFlag As String, it3sResponse As String)

        Dim it3sUniqueNo As Integer
        Dim strit3sUniqueNo As String

        strit3sUniqueNo = it3sResponse.Substring(it3sResponse.IndexOf(DataDelimiter) + 1)
        strit3sUniqueNo = strit3sUniqueNo.Substring(strit3sUniqueNo.IndexOf(DataDelimiter) + 1)
        If strit3sUniqueNo.IndexOf(DataDelimiter) > 0 Then
            strit3sUniqueNo = strit3sUniqueNo.Substring(0, strit3sUniqueNo.IndexOf(DataDelimiter))
        Else
            strit3sUniqueNo = "0"
        End If

        If IsNumeric(strit3sUniqueNo) Then
            it3sUniqueNo = CInt(strit3sUniqueNo)
        Else
            it3sUniqueNo = 0
        End If

        Dim cn As New SqlConnection(strConnection)
        Dim strSQL As String
        Dim cmd As SqlCommand

        strSQL = "INSERT INTO " & IT3sResponseTable & " (TestRun, it3sPeriod, it3sSubmissionNo, it3sUniqueNo, it3sResponse) " &
                        "VALUES ('" & TestFlag & "', '" & it3sSubmissionPeriod & "', " & it3sSubmissionNo & ", " & it3sUniqueNo & ", '" & it3sResponse & "')"
        cmd = New SqlCommand(strSQL, cn)
        Try
            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()
        Catch ex As Exception
        End Try

    End Sub

    Private Function GetResponse(it3sSubmissionPeriod As String, it3sSubmissionNo As Integer, TestFlag As String, it3sUniqueNo As String) As String

        Dim strResponse As String = ""

        Dim strSQL As String = "SELECT it3sResponse " &
                               "FROM " & IT3sResponseTable & " " &
                               "WHERE it3sPeriod = '" & it3sSubmissionPeriod & "' " &
                                        "AND it3sSubmissionNo = " & it3sSubmissionNo & " " &
                                        "AND TestRun = '" & TestFlag & "' " &
                                        "AND it3sUniqueNo = " & it3sUniqueNo

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)

        cn.Open()
        drSQL = cmd.ExecuteReader()
        Try
            drSQL.Read()
            strResponse = drSQL.Item(0)
        Catch ex As Exception
        End Try
        cn.Close()

        Return strResponse

    End Function

    Private Sub UpdateTrailerDatabaseRecord(ByVal hdrId As Integer, ByVal strTrailer As String)

        Dim cn As New SqlConnection(strConnection)
        Dim strSQL As String
        Dim cmd As SqlCommand

        strSQL = "UPDATE " & IT3sHeaderTable & " " &
                 "SET it3sRecord = '" & strTrailer & "' " &
                 "WHERE hdrId = " & hdrId
        cmd = New SqlCommand(strSQL, cn)
        Try
            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()
        Catch ex As Exception
        End Try

    End Sub

    Private Function BuildFileName() As String

        Dim tmpFileName As String

        If SARSForm.OutputPath.Substring(SARSForm.OutputPath.Length) = "\" Then
            tmpFileName = SARSForm.OutputPath
        Else
            tmpFileName = SARSForm.OutputPath & "\"
        End If

        tmpFileName = tmpFileName & FileDataType & FileNameDelimiter &
                      FileLayoutVersion & FileNameDelimiter &
                      InstitutionTaxReference & FileNameDelimiter &
                      SARSForm.txtReference.Text &
                      FixedLengthString(FileSeq, 3, "Right", "0") & FileNameDelimiter &
                      Now.Year.ToString &
                      FixedLengthString(Now.Month.ToString, 2, "Right", "0") &
                      FixedLengthString(Now.Day.ToString, 2, "Right", "0") &
                      "T" &
                      FixedLengthString(Now.Hour.ToString, 2, "Right", "0") &
                      FixedLengthString(Now.Minute.ToString, 2, "Right", "0") &
                      FixedLengthString(Now.Second.ToString, 2, "Right", "0") &
                      ".txt"

        ReDim Preserve ExportFiles(1, ExportFiles.GetUpperBound(1) + 1)
        ExportFiles(0, ExportFiles.GetUpperBound(1)) = tmpFileName

        Return tmpFileName
    End Function

    'Private Sub WriteToFile(ByVal SARSFile As String, ByVal StringToWrite As String, Optional ByVal IncrementCount As Boolean = True)

    '    If StringToWrite <> "" Then
    '        Dim SARSWriter As New System.IO.StreamWriter(SARSFile, True)
    '        SARSWriter.WriteLine(StringToWrite)
    '        SARSWriter.Close()
    '        RecSeq = RecSeq + 1
    '        SARSForm.FileNo = FileSeq
    '        SARSForm.RecordNo = RecSeq
    '        If IncrementCount Then
    '            TotalRecords = TotalRecords + 1
    '        End If
    '    End If
    'End Sub

    Public Sub GetExistingDetails(SubmissionPeriod As String, SubmissionNo As Integer, TestFlag As String)

        'Dim strDBRecord As String
        Dim strRef As String = ""

        Dim strSQL As String = "SELECT it3sReference, it3sFileNo, it3sRecord " &
                               "FROM   tblit3sHeader " &
                               "WHERE  it3sPeriod = '" & SubmissionPeriod & "' " &
                                            "AND it3sSubmissionNo = " & SubmissionNo & " " &
                                            "AND TestRun = '" & TestFlag & "' " &
                               "ORDER BY hdrId ASC"

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand

        cmd = New SqlCommand(strSQL, cn)
        Try
            cn.Open()
            drSQL = cmd.ExecuteReader()
            While drSQL.Read()
                Select Case drSQL.Item(2).ToString.Substring(0, 4)
                    Case "H|GH"
                        strRef = drSQL.Item(0)
                        FileSeq = drSQL.Item(1)
                        If ExportFiles.GetUpperBound(1) < FileSeq - 1 Then
                            ReDim Preserve ExportFiles(1, ExportFiles.GetUpperBound(1) + 1)
                        End If
                        If SARSForm.OutputPath.Substring(SARSForm.OutputPath.Length - 1) = "\" Then
                            ExportFiles(0, FileSeq - 1) = SARSForm.OutputPath & GetColumnFromRecord(5, drSQL.Item(2), DataDelimiter) & ".txt"
                        Else
                            ExportFiles(0, FileSeq - 1) = SARSForm.OutputPath & "\" & GetColumnFromRecord(5, drSQL.Item(2), DataDelimiter) & ".txt"
                        End If
                    Case Else
                        If drSQL.Item(2).ToString.Substring(0, 1) = "T" Then
                            ExportFiles(1, FileSeq - 1) = GetColumnFromRecord(2, drSQL.Item(2), DataDelimiter)
                        End If
                End Select
            End While
            cn.Close()
            SARSForm.Reference = strRef
        Catch ex As Exception
            SARSForm.OnlyCreateFiles = False
            MessageBox.Show("There does not appear to be any existing data in the database - please check your facts", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub WriteToDatabase(ByVal tblDestination As String, ByVal StringToWrite As String, Optional ByVal IncrementCount As Boolean = True)

        Dim cn As New SqlConnection(strConnection)
        Dim strSQL As String
        Dim cmd As SqlCommand

        strSQL = "INSERT INTO " & tblDestination & "(TestRun, it3sPeriod, it3sSubmissionNo, it3sReference, it3sFileNo, it3sRecord) " &
                     "VALUES('" & SARSForm.TestFile & "', '" & it3sPeriod & "', " & it3sSubmissionNo & ", '" & SARSForm.txtReference.Text & "', " & FileSeq & ", '" & StringToWrite.Replace("'", "") & "')"
        cmd = New SqlCommand(strSQL, cn)
        Try
            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()
            RecSeq = RecSeq + 1
            SARSForm.FileNo = FileSeq
            SARSForm.RecordNo = RecSeq
            If IncrementCount Then
                TotalRecords = TotalRecords + 1
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub WriteExceptionToDatabase(ByVal tblDestination As String, ByVal strCIF As String, ByVal strCategory As String, ByVal StringToWrite As String)

        Dim cn As New SqlConnection(strConnection)
        Dim strSQL As String
        Dim cmd As SqlCommand

        strSQL = "INSERT INTO " & tblDestination & "(TestRun, it3sPeriod, it3sSubmissionNo, it3sReference, it3sCIF, it3sExceptionCategory, it3sExceptionDescription) " &
                     "VALUES('" & SARSForm.TestFile & "', '" & it3sPeriod & "', " & it3sSubmissionNo & ", '" & SARSForm.txtReference.Text & "', '" & strCIF & "', '" & strCategory & "', '" & StringToWrite & "')"
        cmd = New SqlCommand(strSQL, cn)
        Try
            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()
        Catch ex As Exception
        End Try
    End Sub

    'Private Sub WriteTrailerWithHashTotal(ByVal curFileNo As Integer, ByVal strFileName As String, ByVal recId As Integer, ByVal strTrailer As String)

    '    If Not SARSForm.OnlyCreateFiles Then

    '        Dim md5code As String

    '        Dim md5 As MD5CryptoServiceProvider = New MD5CryptoServiceProvider
    '        Dim f As FileStream = New FileStream(strFileName, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
    '        'f = New FileStream(OpenFileDialog1.FileName, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
    '        md5.ComputeHash(f)
    '        'Dim ObjFSO As Object = CreateObject("Scripting.FileSystemObject")
    '        'Dim objFile = ObjFSO.GetFile(OpenFileDialog1.FileName)

    '        Dim hash As Byte() = md5.Hash
    '        Dim buff As StringBuilder = New StringBuilder
    '        Dim hashByte As Byte
    '        For Each hashByte In hash
    '            buff.Append(String.Format("{0:X2}", hashByte))
    '        Next

    '        md5code = buff.ToString()

    '        f.Close()

    '        strTrailer = strTrailer.Replace("HASH_TOTAL_" & curFileNo, md5code)

    '        UpdateTrailerDatabaseRecord(recId, strTrailer)

    '    End If

    '    Dim ExportFile As New StreamWriter(strFileName, True)
    '    ExportFile.WriteLine(strTrailer)
    '    ExportFile.Close()

    'End Sub

    Private Sub WriteTrailerWithHashTotal(ByVal curFileNo As Integer, ByVal strFileName As String, ByVal recId As Integer, ByVal strTrailer As String)

        If Not SARSForm.OnlyCreateFiles Then

            'strTrailer = strTrailer.Replace("HASH_TOTAL_" & curFileNo, GetMD5HashByLine(strFileName))
            strTrailer = strTrailer.Replace("HASH_TOTAL_" & curFileNo, GetMD5Hash(strFileName))

            UpdateTrailerDatabaseRecord(recId, strTrailer)

        End If

        Dim ExportFile As New StreamWriter(strFileName, True)
        ExportFile.WriteLine(strTrailer)
        ExportFile.Close()

    End Sub

    Private Function GetMD5Hash(ByVal strFileName As String) As String

        Dim fileContents As String

        Try
            Dim ReadFile As New StreamReader(strFileName)
            fileContents = ReadFile.ReadToEnd
            ReadFile.Close()
        Catch ex As Exception
            fileContents = ""
        End Try

        fileContents = fileContents.Replace(vbCrLf, "")

        Return MD5Hasher.HashTotal(fileContents)

    End Function

    'Private Function GetMD5HashByLine(ByVal strFileName As String) As String

    '    'Dim strLine As String
    '    Dim myLongString As String = ""
    '    'Dim tmpByteValue As Byte()
    '    Dim ByteValue As Byte() 'These are both byte arrays
    '    'Dim ByteHash() As Byte 'I was just demonstrating different ways to declare them
    '    Dim Output As String
    '    Dim Provider = New MD5CryptoServiceProvider()
    '    'Dim i As Integer
    '    'Dim j As Integer

    '    'Dim FileStream As New IO.FileStream(Input, IO.FileMode.Open, IO.FileAccess.Read)
    '    'Dim BinaryReader As New IO.BinaryReader(FileStream)

    '    Dim ReadFile As New StreamReader(strFileName)

    '    While ReadFile.Peek <> -1
    '        myLongString = myLongString & ReadFile.ReadLine()
    '    End While

    '    ReadFile.Close()

    '    ByteValue = System.Text.Encoding.ASCII.GetBytes(myLongString)
    '    ByteValue = Provider.ComputeHash(ByteValue)
    '    For Each b As Byte In ByteValue
    '        Output += b.ToString("x2")
    '    Next

    '    Return Output
    'End Function

    'Private Function GetMD5HashByChar(ByVal strFileName As String) As String

    '    Dim myLongString As String = ""
    '    Dim tmpByteValue As Byte()
    '    Dim ByteValue As Byte() 'These are both byte arrays
    '    Dim ByteHash() As Byte 'I was just demonstrating different ways to declare them
    '    Dim Output As String
    '    Dim Provider = New MD5CryptoServiceProvider()
    '    Dim i As Integer
    '    Dim j As Integer

    '    Dim FileStream As New IO.FileStream(strFileName, IO.FileMode.Open, IO.FileAccess.Read)
    '    Dim BinaryReader As New IO.BinaryReader(FileStream)

    '    tmpByteValue = BinaryReader.ReadBytes(FileStream.Length)
    '    j = 0
    '    For i = 0 To tmpByteValue.Length - 1
    '        If tmpByteValue(i) <> 10 And tmpByteValue(i) <> 13 Then
    '            ReDim Preserve ByteValue(j)
    '            ByteValue(j) = tmpByteValue(i)
    '            j += 1
    '        End If
    '    Next
    '    ByteHash = Provider.ComputeHash(ByteValue)
    '    FileStream.Close()
    '    BinaryReader.Close()
    '    For i = 0 To ByteHash.Length - 1
    '        Output += ByteHash(i).ToString("x").PadLeft(2, "0")
    '    Next

    '    Provider.Clear()
    '    ByteValue = Nothing
    '    Return Output
    'End Function

    'Private Function tmpGetMD5Hash(ByVal Input As String, Optional ByVal IsFile As Boolean = False) As String

    '    Dim ByteValue As Byte() 'These are both byte arrays
    '    Dim ByteHash() As Byte 'I was just demonstrating different ways to declare them
    '    Dim Output As String
    '    Dim Provider = New MD5CryptoServiceProvider()
    '    Dim i As Integer

    '    If IsFile Then
    '        If IO.File.Exists(Input) Then
    '            Dim FileStream As New IO.FileStream(Input, IO.FileMode.Open, IO.FileAccess.Read)
    '            Dim BinaryReader As New IO.BinaryReader(FileStream)
    '            ByteValue = BinaryReader.ReadBytes(FileStream.Length)
    '            ByteHash = Provider.ComputeHash(ByteValue)
    '            FileStream.Close()
    '            BinaryReader.Close()
    '            For i = 0 To ByteHash.Length - 1
    '                Output += ByteHash(i).ToString("x").PadLeft(2, "0")
    '            Next
    '        Else : Return "File Not Found"
    '        End If
    '    Else
    '        ByteValue = System.Text.Encoding.ASCII.GetBytes(Input)
    '        ByteValue = Provider.ComputeHash(ByteValue)
    '        For Each b As Byte In ByteValue
    '            Output += b.ToString("x2")
    '        Next
    '    End If
    '    Provider.Clear()
    '    ByteValue = Nothing
    '    Return Output
    'End Function

    Private Function GetColumnFromRecord(ByVal colNo As Integer, ByVal strRecord As String, ByVal strDelimiter As String) As String

        Dim col() As String = strRecord.Split(strDelimiter)

        Try
            Return col(colNo - 1)
        Catch ex As Exception
            Return ""
        End Try

    End Function

    'Private Function GetColumnFromRecord(ByVal colNo As Integer, ByVal strRecord As String, ByVal strDelimiter As String) As String

    '    Dim i As Integer
    '    Dim DelimiterCount As Integer = 0
    '    Dim idxStart As Integer = -1
    '    Dim idxEnd As Integer = -1

    '    Try
    '        i = -1
    '        While i < strRecord.Length And DelimiterCount < colNo
    '            i += 1
    '            If strRecord.Substring(i, 1) = strDelimiter Then
    '                DelimiterCount += 1
    '                If DelimiterCount = colNo - 1 Then
    '                    idxStart = i
    '                End If
    '                If DelimiterCount = colNo Then
    '                    idxEnd = i
    '                End If
    '            End If
    '        End While
    '        If idxStart < 0 Then
    '            Return ""
    '        ElseIf idxEnd < 0 Then
    '            Return strRecord.Substring(idxStart + 1).Replace("-", FileNameDelimiter)
    '        Else
    '            Return strRecord.Substring(idxStart + 1, idxEnd - idxStart - 1).Replace("-", FileNameDelimiter)
    '        End If
    '    Catch ex As Exception
    '        Return ""
    '    End Try
    'End Function

    'Shared Sub Main(ByVal args() As String)
    '    Dim [source] As String = "Hello World!"
    '    Using md5Hash As MD5 = MD5.Create()

    '        Dim hash As String = GetMd5Hash(md5Hash, source)

    '        Console.WriteLine("The MD5 hash of " + source + " is: " + hash + ".")

    '        Console.WriteLine("Verifying the hash...")

    '        If VerifyMd5Hash(md5Hash, [source], hash) Then
    '            Console.WriteLine("The hashes are the same.")
    '        Else
    '            Console.WriteLine("The hashes are not same.")
    '        End If
    '    End Using
    'End Sub 'Main



    'Shared Function GetMd5Hash(ByVal md5Hash As MD5, ByVal input As String) As String

    '    ' Convert the input string to a byte array and compute the hash. 
    '    Dim data As Byte() = md5Hash.ComputeHash(Encoding.UTF8.GetBytes(input))

    '    ' Create a new Stringbuilder to collect the bytes 
    '    ' and create a string. 
    '    Dim sBuilder As New StringBuilder()

    '    ' Loop through each byte of the hashed data  
    '    ' and format each one as a hexadecimal string. 
    '    Dim i As Integer
    '    For i = 0 To data.Length - 1
    '        sBuilder.Append(data(i).ToString("x2"))
    '    Next i

    '    ' Return the hexadecimal string. 
    '    Return sBuilder.ToString()

    'End Function 'GetMd5Hash


    '' Verify a hash against a string. 
    'Shared Function VerifyMd5Hash(ByVal md5Hash As MD5, ByVal input As String, ByVal hash As String) As Boolean
    '    ' Hash the input. 
    '    Dim hashOfInput As String = GetMd5Hash(md5Hash, input)

    '    ' Create a StringComparer an compare the hashes. 
    '    Dim comparer As StringComparer = StringComparer.OrdinalIgnoreCase

    '    If 0 = comparer.Compare(hashOfInput, hash) Then
    '        Return True
    '    Else
    '        Return False
    '    End If

    'End Function 'VerifyMd5Hash

End Module