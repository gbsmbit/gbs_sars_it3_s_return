Imports System.IO
Imports System.Data.SqlClient

Public Class SARSIT3Form
    Inherits System.Windows.Forms.Form
    Private TestValue As String = "T"
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents lblDatabase As System.Windows.Forms.Label
    Friend WithEvents txtPeriodLength As System.Windows.Forms.ComboBox
    Friend WithEvents chkCreateFilesFromExistingData As System.Windows.Forms.CheckBox
    Friend WithEvents chkSkipInvalidReferences As System.Windows.Forms.CheckBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lblSubmissionNo As System.Windows.Forms.Label
    Friend WithEvents chkClearDatabase As System.Windows.Forms.CheckBox
    Friend WithEvents btnImportResponseFile As System.Windows.Forms.Button
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents lblSkipList As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents lblSkipInactive As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents lblSkipReferences As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents lblSkipPersonalInfo As System.Windows.Forms.Label
    Private LiveValue As String = "L"

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub
    Private components As System.ComponentModel.IContainer

    'Required by the Windows Form Designer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnExecute As System.Windows.Forms.Button
    Friend WithEvents txtTest As System.Windows.Forms.TextBox
    Friend WithEvents txtReference As System.Windows.Forms.TextBox
    Friend WithEvents txtPath As System.Windows.Forms.TextBox
    Friend WithEvents Period As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lblFileNo As System.Windows.Forms.Label
    Friend WithEvents lblRecNo As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents lblClientRec As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents lblSkipped As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents lblTimePassed As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents btnPathHelp As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SARSIT3Form))
        Me.txtTest = New System.Windows.Forms.TextBox()
        Me.txtReference = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtPath = New System.Windows.Forms.TextBox()
        Me.btnExecute = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.Period = New System.Windows.Forms.DateTimePicker()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lblFileNo = New System.Windows.Forms.Label()
        Me.lblRecNo = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblClientRec = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lblSkipped = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.lblTimePassed = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.btnPathHelp = New System.Windows.Forms.Button()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.lblDatabase = New System.Windows.Forms.Label()
        Me.txtPeriodLength = New System.Windows.Forms.ComboBox()
        Me.chkCreateFilesFromExistingData = New System.Windows.Forms.CheckBox()
        Me.chkSkipInvalidReferences = New System.Windows.Forms.CheckBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblSubmissionNo = New System.Windows.Forms.Label()
        Me.chkClearDatabase = New System.Windows.Forms.CheckBox()
        Me.btnImportResponseFile = New System.Windows.Forms.Button()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.lblSkipList = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.lblSkipInactive = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.lblSkipReferences = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.lblSkipPersonalInfo = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'txtTest
        '
        Me.txtTest.Location = New System.Drawing.Point(822, 181)
        Me.txtTest.MaxLength = 1
        Me.txtTest.Name = "txtTest"
        Me.txtTest.Size = New System.Drawing.Size(66, 26)
        Me.txtTest.TabIndex = 2
        '
        'txtReference
        '
        Me.txtReference.Location = New System.Drawing.Point(277, 181)
        Me.txtReference.MaxLength = 14
        Me.txtReference.Name = "txtReference"
        Me.txtReference.Size = New System.Drawing.Size(301, 26)
        Me.txtReference.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(64, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(160, 34)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Period End Date:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(602, 186)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(211, 33)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Test File [(T)est/(L)ive]:"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(64, 181)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(160, 34)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Unique Reference:"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(64, 240)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(160, 33)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Output Path:"
        '
        'txtPath
        '
        Me.txtPath.Location = New System.Drawing.Point(277, 240)
        Me.txtPath.Name = "txtPath"
        Me.txtPath.Size = New System.Drawing.Size(611, 26)
        Me.txtPath.TabIndex = 5
        '
        'btnExecute
        '
        Me.btnExecute.Location = New System.Drawing.Point(606, 409)
        Me.btnExecute.Name = "btnExecute"
        Me.btnExecute.Size = New System.Drawing.Size(120, 34)
        Me.btnExecute.TabIndex = 6
        Me.btnExecute.Text = "Execute"
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(773, 409)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(120, 34)
        Me.btnExit.TabIndex = 7
        Me.btnExit.Text = "Exit"
        '
        'Period
        '
        Me.Period.Location = New System.Drawing.Point(277, 64)
        Me.Period.Name = "Period"
        Me.Period.Size = New System.Drawing.Size(256, 26)
        Me.Period.TabIndex = 1
        Me.Period.Value = New Date(2014, 2, 28, 0, 0, 0, 0)
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(61, 360)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(102, 33)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "File No.:"
        '
        'lblFileNo
        '
        Me.lblFileNo.Location = New System.Drawing.Point(278, 360)
        Me.lblFileNo.Name = "lblFileNo"
        Me.lblFileNo.Size = New System.Drawing.Size(160, 33)
        Me.lblFileNo.TabIndex = 13
        '
        'lblRecNo
        '
        Me.lblRecNo.Location = New System.Drawing.Point(278, 395)
        Me.lblRecNo.Name = "lblRecNo"
        Me.lblRecNo.Size = New System.Drawing.Size(160, 33)
        Me.lblRecNo.TabIndex = 15
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(61, 395)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(102, 33)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "Record No.:"
        '
        'lblClientRec
        '
        Me.lblClientRec.Location = New System.Drawing.Point(278, 430)
        Me.lblClientRec.Name = "lblClientRec"
        Me.lblClientRec.Size = New System.Drawing.Size(160, 33)
        Me.lblClientRec.TabIndex = 17
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(61, 430)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(192, 33)
        Me.Label8.TabIndex = 16
        Me.Label8.Text = "Total Client Records:"
        '
        'lblSkipped
        '
        Me.lblSkipped.Location = New System.Drawing.Point(278, 465)
        Me.lblSkipped.Name = "lblSkipped"
        Me.lblSkipped.Size = New System.Drawing.Size(160, 33)
        Me.lblSkipped.TabIndex = 19
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(61, 465)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(205, 33)
        Me.Label9.TabIndex = 18
        Me.Label9.Text = "Clients skipped:"
        '
        'lblTimePassed
        '
        Me.lblTimePassed.Location = New System.Drawing.Point(662, 465)
        Me.lblTimePassed.Name = "lblTimePassed"
        Me.lblTimePassed.Size = New System.Drawing.Size(218, 33)
        Me.lblTimePassed.TabIndex = 21
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(509, 465)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(128, 33)
        Me.Label10.TabIndex = 20
        Me.Label10.Text = "Time Elapsed:"
        '
        'Timer1
        '
        '
        'btnPathHelp
        '
        Me.btnPathHelp.BackgroundImage = CType(resources.GetObject("btnPathHelp.BackgroundImage"), System.Drawing.Image)
        Me.btnPathHelp.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnPathHelp.Location = New System.Drawing.Point(832, 64)
        Me.btnPathHelp.Name = "btnPathHelp"
        Me.btnPathHelp.Size = New System.Drawing.Size(56, 47)
        Me.btnPathHelp.TabIndex = 8
        Me.btnPathHelp.UseVisualStyleBackColor = True
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(64, 115)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(160, 34)
        Me.Label11.TabIndex = 25
        Me.Label11.Text = "Period Length:"
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(59, 13)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(259, 34)
        Me.Label12.TabIndex = 26
        Me.Label12.Text = "Database Connection String:"
        '
        'lblDatabase
        '
        Me.lblDatabase.Location = New System.Drawing.Point(291, 13)
        Me.lblDatabase.Name = "lblDatabase"
        Me.lblDatabase.Size = New System.Drawing.Size(619, 47)
        Me.lblDatabase.TabIndex = 27
        '
        'txtPeriodLength
        '
        Me.txtPeriodLength.FormattingEnabled = True
        Me.txtPeriodLength.Items.AddRange(New Object() {"6", "12"})
        Me.txtPeriodLength.Location = New System.Drawing.Point(277, 111)
        Me.txtPeriodLength.Name = "txtPeriodLength"
        Me.txtPeriodLength.Size = New System.Drawing.Size(77, 28)
        Me.txtPeriodLength.TabIndex = 28
        Me.txtPeriodLength.Text = "12"
        '
        'chkCreateFilesFromExistingData
        '
        Me.chkCreateFilesFromExistingData.Location = New System.Drawing.Point(514, 327)
        Me.chkCreateFilesFromExistingData.Name = "chkCreateFilesFromExistingData"
        Me.chkCreateFilesFromExistingData.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkCreateFilesFromExistingData.Size = New System.Drawing.Size(374, 25)
        Me.chkCreateFilesFromExistingData.TabIndex = 29
        Me.chkCreateFilesFromExistingData.Text = "Create submission files from existing data"
        Me.chkCreateFilesFromExistingData.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkCreateFilesFromExistingData.UseVisualStyleBackColor = True
        '
        'chkSkipInvalidReferences
        '
        Me.chkSkipInvalidReferences.Location = New System.Drawing.Point(146, 294)
        Me.chkSkipInvalidReferences.Name = "chkSkipInvalidReferences"
        Me.chkSkipInvalidReferences.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkSkipInvalidReferences.Size = New System.Drawing.Size(742, 25)
        Me.chkSkipInvalidReferences.TabIndex = 30
        Me.chkSkipInvalidReferences.Text = "Skip clients where ID/Company Registration No. and Tax Reference fail modulus che" &
    "ck"
        Me.chkSkipInvalidReferences.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkSkipInvalidReferences.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(59, 329)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(165, 33)
        Me.Label6.TabIndex = 31
        Me.Label6.Text = "Submission No.:"
        '
        'lblSubmissionNo
        '
        Me.lblSubmissionNo.Location = New System.Drawing.Point(278, 329)
        Me.lblSubmissionNo.Name = "lblSubmissionNo"
        Me.lblSubmissionNo.Size = New System.Drawing.Size(160, 33)
        Me.lblSubmissionNo.TabIndex = 32
        '
        'chkClearDatabase
        '
        Me.chkClearDatabase.Location = New System.Drawing.Point(514, 358)
        Me.chkClearDatabase.Name = "chkClearDatabase"
        Me.chkClearDatabase.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkClearDatabase.Size = New System.Drawing.Size(374, 25)
        Me.chkClearDatabase.TabIndex = 33
        Me.chkClearDatabase.Text = "Clear all data for specified period"
        Me.chkClearDatabase.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkClearDatabase.UseVisualStyleBackColor = True
        '
        'btnImportResponseFile
        '
        Me.btnImportResponseFile.Location = New System.Drawing.Point(606, 512)
        Me.btnImportResponseFile.Name = "btnImportResponseFile"
        Me.btnImportResponseFile.Size = New System.Drawing.Size(287, 33)
        Me.btnImportResponseFile.TabIndex = 34
        Me.btnImportResponseFile.Text = "Import Response File"
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(104, 560)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(165, 33)
        Me.Label16.TabIndex = 49
        Me.Label16.Text = "Skip List:"
        '
        'lblSkipList
        '
        Me.lblSkipList.Location = New System.Drawing.Point(282, 560)
        Me.lblSkipList.Name = "lblSkipList"
        Me.lblSkipList.Size = New System.Drawing.Size(160, 33)
        Me.lblSkipList.TabIndex = 50
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(104, 539)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(165, 34)
        Me.Label15.TabIndex = 47
        Me.Label15.Text = "Inactive:"
        '
        'lblSkipInactive
        '
        Me.lblSkipInactive.Location = New System.Drawing.Point(282, 539)
        Me.lblSkipInactive.Name = "lblSkipInactive"
        Me.lblSkipInactive.Size = New System.Drawing.Size(160, 34)
        Me.lblSkipInactive.TabIndex = 48
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(104, 519)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(165, 33)
        Me.Label14.TabIndex = 45
        Me.Label14.Text = "References:"
        '
        'lblSkipReferences
        '
        Me.lblSkipReferences.Location = New System.Drawing.Point(282, 519)
        Me.lblSkipReferences.Name = "lblSkipReferences"
        Me.lblSkipReferences.Size = New System.Drawing.Size(160, 33)
        Me.lblSkipReferences.TabIndex = 46
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(104, 498)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(165, 34)
        Me.Label13.TabIndex = 43
        Me.Label13.Text = "Personal Info:"
        '
        'lblSkipPersonalInfo
        '
        Me.lblSkipPersonalInfo.Location = New System.Drawing.Point(282, 498)
        Me.lblSkipPersonalInfo.Name = "lblSkipPersonalInfo"
        Me.lblSkipPersonalInfo.Size = New System.Drawing.Size(160, 34)
        Me.lblSkipPersonalInfo.TabIndex = 44
        '
        'SARSIT3Form
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 19)
        Me.ClientSize = New System.Drawing.Size(945, 608)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.lblSkipList)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.lblSkipInactive)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.lblSkipReferences)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.lblSkipPersonalInfo)
        Me.Controls.Add(Me.btnImportResponseFile)
        Me.Controls.Add(Me.chkClearDatabase)
        Me.Controls.Add(Me.lblSubmissionNo)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.chkSkipInvalidReferences)
        Me.Controls.Add(Me.chkCreateFilesFromExistingData)
        Me.Controls.Add(Me.txtPeriodLength)
        Me.Controls.Add(Me.lblDatabase)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.btnPathHelp)
        Me.Controls.Add(Me.lblTimePassed)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.lblSkipped)
        Me.Controls.Add(Me.lblClientRec)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.lblRecNo)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.lblFileNo)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Period)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnExecute)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtPath)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtReference)
        Me.Controls.Add(Me.txtTest)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "SARSIT3Form"
        Me.Text = "SARS Third Party Data Submission IT3(s)"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Function MyDateString(ByVal ADate As Date) As String

        MyDateString = CStr(ADate)
        MyDateString = MyDateString.Substring(0, 4) & MyDateString.Substring(5, 2) & MyDateString.Substring(8, 2)
    End Function

    Public ReadOnly Property PeriodEnd() As String
        Get
            Return MyDateString(CDate("#" & Period.Value.Month & "/" & Period.Value.Day & "/" & Period.Value.Year & "#"))
        End Get
    End Property

    Public ReadOnly Property PeriodStart(ByVal PeriodLength As Byte) As String
        Get
            'Return MyDateString(DateAdd( DateInterval.Day, 1, DateAdd(DateInterval.Year, -1, CDate("#" & Period.Value.Month & "/" & Period.Value.Day & "/" & Period.Value.Year & "#"))))
            'Return MyDateString(DateAdd(DateInterval.Day, 1, DateAdd(DateInterval.Month, PeriodLength * -1, CDate("#" & Period.Value.Month & "/" & Period.Value.Day & "/" & Period.Value.Year & "#"))))
            Return MyDateString(DateSerial(Period.Value.Year, Period.Value.Month - PeriodLength + 1, 1))
        End Get
    End Property

    Public Property OnlyCreateFiles() As Boolean
        Get
            Return chkCreateFilesFromExistingData.Checked
        End Get
        Set(ByVal Value As Boolean)
            chkCreateFilesFromExistingData.Checked = Value
        End Set
    End Property

    Public Property SkipInvalidReferences() As Boolean
        Get
            Return chkSkipInvalidReferences.Checked
        End Get
        Set(ByVal Value As Boolean)
            chkSkipInvalidReferences.Checked = Value
        End Set
    End Property

    Public ReadOnly Property TestFile() As String
        Get
            Return txtTest.Text
        End Get
    End Property

    'Public ReadOnly Property FileSeqStart() As Integer
    '    Get
    '        Return CInt(txtSeq.Text)
    '    End Get
    'End Property

    Public Property Reference() As String
        Get
            Return txtReference.Text
        End Get
        Set(ByVal Value As String)
            txtReference.Text = Value
        End Set
    End Property

    Public Property OutputPath() As String
        Get
            Return txtPath.Text
        End Get
        Set(ByVal Value As String)
            txtPath.Text = Value
        End Set
    End Property

    Public WriteOnly Property FileNo() As String
        Set(ByVal Value As String)
            lblFileNo.Text = Value
        End Set
    End Property

    Public WriteOnly Property RecordNo() As String
        Set(ByVal Value As String)
            lblRecNo.Text = Value
        End Set
    End Property

    Public WriteOnly Property ClientRecords() As String
        Set(ByVal Value As String)
            lblClientRec.Text = Value
        End Set
    End Property

    Public WriteOnly Property ClientsSkipped() As String
        Set(ByVal Value As String)
            lblSkipped.Text = Value
        End Set
    End Property

    Public WriteOnly Property ClientsSkippedDueToPersonalInfo() As String
        Set(ByVal Value As String)
            lblSkipPersonalInfo.Text = Value
        End Set
    End Property

    Public WriteOnly Property ClientsSkippedDueToReferences() As String
        Set(ByVal Value As String)
            lblSkipReferences.Text = Value
        End Set
    End Property

    Public WriteOnly Property ClientsSkippedDueInactivity() As String
        Set(ByVal Value As String)
            lblSkipInactive.Text = Value
        End Set
    End Property

    Public WriteOnly Property ClientsSkippedDueToList() As String
        Set(ByVal Value As String)
            lblSkipList.Text = Value
        End Set
    End Property

    Public WriteOnly Property TimeElapsed() As String
        Set(ByVal Value As String)
            lblTimePassed.Text = Value
        End Set
    End Property

    Public WriteOnly Property SetButton() As String
        Set(ByVal Value As String)
            btnExit.Text = Value
        End Set
    End Property

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click

        If Exporting Then
            Abort = True
        Else
            Me.Close()
        End If
    End Sub

    Private Sub btnExecute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExecute.Click

        Dim OKToContinue As Boolean = False

        If chkCreateFilesFromExistingData.Checked Then
            If MessageBox.Show("Any existing files with the same references will be overwritten.  Continue?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                OKToContinue = True
            End If
        Else
            If chkClearDatabase.Checked Then
                'If it3sSubmissionNo > 0 And MessageBox.Show("Multiple submissions exist for this period.  Delete all?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                '    OKToContinue = True
                'ElseIf MessageBox.Show("Delete original submission for this period?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                '    OKToContinue = True
                'End If
                If MessageBox.Show("Delete all submission records for this period?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    OKToContinue = True
                End If
            Else
                OKToContinue = True
            End If
        End If
        If OKToContinue Then

#If Not DEBUG Then
            Timer1.Enabled = True
#End If

            ClearUniqueNumbers()
            If Not chkCreateFilesFromExistingData.Checked Then
                If chkClearDatabase.Checked Then
                    ClearDatabase(it3sPeriod)
                    it3sSubmissionNo = GetSubmissionNo(it3sPeriod)
                    lblSubmissionNo.Text = it3sSubmissionNo
                End If
                StartTime = Now
                CreateExportData()
            Else
                StartTime = Now
            End If
            CreateExportFiles(it3sPeriod, it3sSubmissionNo, txtTest.Text)
            Timer1.Enabled = False
            If Abort Then
                MessageBox.Show("Processing aborted by user", "Done", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                MessageBox.Show("Processing successfully completed", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If
    End Sub

    Private Sub SARSIT3Form_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.Text = "SARS Third Party Data Submission IT3(s) - V" & My.Settings.AppVersion & " (BRS " & My.Settings.SARS_BRS & ")"

        If Month(Today) > 8 Then
            Period.Value = DateSerial(Year(Today), 8, 31)
            txtPeriodLength.Text = "6"
        Else
            Period.Value = DateAdd(DateInterval.Day, -1, DateSerial(Year(Today), 3, 1))
            txtPeriodLength.Text = "12"
        End If
        lblDatabase.Text = strConnection.Substring(0, strConnection.LastIndexOf("=") + 1) & "**********"
        txtTest.Text = TestValue
        txtPath.Text = DataFileLocationTest
        SkipInvalidReferences = SkipInvalidRefs
        Refresh()
        MessageBox.Show("Note:  For this application to produce the correct output files, it is necessary to first run the Tax Certificate application for the required period.", _
                        "Reminder", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub SetUniqueReference()

        it3sPeriod = Period.Value.Year & FixedLengthString(Period.Value.Month, 2, "Right", "0") & FixedLengthString(Period.Value.Day, 2, "Right", "0")
        txtReference.Text = "GBS" & it3sPeriod & _
                             FixedLengthString(TimeOfDay.Minute.ToString, 3, "Right", "0") & FixedLengthString(TimeOfDay.Second.ToString, 3, "Right", "0")
        SetSubmissionNo()
    End Sub

    Private Function GetSubmissionNo(SubmissionPeriod As String) As Integer

        Dim strSQL As String = "SELECT MAX(it3sSubmissionNo) " & _
                               "FROM " & IT3sHeaderTable & " " & _
                               "WHERE it3sPeriod = '" & SubmissionPeriod & "' " & _
                                        "AND TestRun = '" & txtTest.Text & "'"

        Dim cn As SqlConnection = New SqlConnection(strConnection)
        Dim drSQL As SqlDataReader
        Dim cmd As SqlCommand = New SqlCommand(strSQL, cn)
        Dim retValue As Integer

        Try
            cn.Open()
            drSQL = cmd.ExecuteReader()
            If drSQL.Read() Then
                If chkCreateFilesFromExistingData.Checked Then
                    retValue = drSQL.Item(0)
                Else
                    retValue = drSQL.Item(0) + 1
                End If
            End If
        Catch ex As Exception
            retValue = 0
        End Try
        cn.Close()
        Return retValue
    End Function

    Private Sub SetSubmissionNo()

        it3sSubmissionNo = GetSubmissionNo(it3sPeriod)
        lblSubmissionNo.Text = it3sSubmissionNo
    End Sub

    Private Sub Period_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Period.ValueChanged

        SetUniqueReference()
    End Sub

    Private Sub txtTest_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTest.TextChanged

        If Not (txtTest.Text.ToUpper = TestValue Or txtTest.Text.ToUpper = LiveValue) Then
            txtTest.Text = ""
            txtPath.Text = ""
        Else
            txtTest.Text = txtTest.Text.ToUpper
            Select Case txtTest.Text.ToUpper
                Case TestValue
                    txtPath.Text = DataFileLocationTest
                Case LiveValue
                    txtPath.Text = DataFileLocation
            End Select
        End If
        SetSubmissionNo()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Dim SomeTime As Date
        Dim SecondsPassed As Integer
        Dim myHours As Integer
        Dim myMinutes As Integer

        SomeTime = Now
        SecondsPassed = DateDiff(DateInterval.Second, StartTime, SomeTime)
        myHours = SecondsPassed \ 3600
        SecondsPassed = SecondsPassed - (myHours * 3600)
        myMinutes = SecondsPassed \ 60
        SecondsPassed = SecondsPassed - (myMinutes * 60)
        lblTimePassed.Text = Format(myHours, "00") & ":" & Format(myMinutes, "00") & ":" & Format(SecondsPassed, "00")
    End Sub

    Private Sub btnPathHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPathHelp.Click

        MessageBox.Show("Default values for all parameters can be specified in the database table mblIT3Parameters, or by saving them in the file " & IniFile, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub chkCreateFilesFromExistingData_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCreateFilesFromExistingData.CheckedChanged

        SetSubmissionNo()
        If chkCreateFilesFromExistingData.Checked Then
            GetExistingDetails(it3sPeriod, it3sSubmissionNo, txtTest.Text)
        End If
    End Sub

    Private Sub txtPeriodLength_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles txtPeriodLength.SelectedIndexChanged

        SetSubmissionNo()
    End Sub

    Private Sub btnImportResponseFile_Click(sender As System.Object, e As System.EventArgs) Handles btnImportResponseFile.Click

        Dim openResponseFile As New OpenFileDialog()
        Dim EffectiveSubmissionNo As Integer

        If chkCreateFilesFromExistingData.Checked Then
            EffectiveSubmissionNo = it3sSubmissionNo
        Else
            EffectiveSubmissionNo = it3sSubmissionNo - 1
        End If

        openResponseFile.InitialDirectory = txtPath.Text.Substring(0, txtPath.Text.LastIndexOf("\"))
        openResponseFile.Filter = "Text files|*.txt"
        openResponseFile.Title = "Select Response File"

        If openResponseFile.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            Try

                Dim Responses As New StreamReader(openResponseFile.FileName)

                While Responses.Peek <> -1
                    InsertResponse(it3sPeriod, EffectiveSubmissionNo, txtTest.Text, Responses.ReadLine())
                End While
                Responses.Close()
                MessageBox.Show("Response file successfully imported", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("An error occurred while reading the response file", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End If
    End Sub

End Class
