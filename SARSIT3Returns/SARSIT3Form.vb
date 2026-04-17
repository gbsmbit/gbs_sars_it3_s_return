Public Class SARSIT3Form
    Inherits System.Windows.Forms.Form
    Private TestValue As String = "Y"

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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
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
        Me.SuspendLayout()
        '
        'txtTest
        '
        Me.txtTest.Location = New System.Drawing.Point(176, 64)
        Me.txtTest.MaxLength = 1
        Me.txtTest.Name = "txtTest"
        Me.txtTest.TabIndex = 1
        Me.txtTest.Text = ""
        '
        'txtReference
        '
        Me.txtReference.Location = New System.Drawing.Point(176, 104)
        Me.txtReference.MaxLength = 14
        Me.txtReference.Name = "txtReference"
        Me.txtReference.TabIndex = 2
        Me.txtReference.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(40, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Period End Date:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(40, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Test File (Y/N):"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(40, 104)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Unique Reference:"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(40, 144)
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Output Path:"
        '
        'txtPath
        '
        Me.txtPath.Location = New System.Drawing.Point(176, 144)
        Me.txtPath.Name = "txtPath"
        Me.txtPath.Size = New System.Drawing.Size(400, 20)
        Me.txtPath.TabIndex = 6
        Me.txtPath.Text = ""
        '
        'btnExecute
        '
        Me.btnExecute.Location = New System.Drawing.Point(392, 200)
        Me.btnExecute.Name = "btnExecute"
        Me.btnExecute.TabIndex = 8
        Me.btnExecute.Text = "Execute"
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(496, 200)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.TabIndex = 9
        Me.btnExit.Text = "Exit"
        '
        'Period
        '
        Me.Period.Location = New System.Drawing.Point(176, 24)
        Me.Period.Name = "Period"
        Me.Period.Size = New System.Drawing.Size(160, 20)
        Me.Period.TabIndex = 11
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(56, 184)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(64, 23)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "File No.:"
        '
        'lblFileNo
        '
        Me.lblFileNo.Location = New System.Drawing.Point(192, 184)
        Me.lblFileNo.Name = "lblFileNo"
        Me.lblFileNo.TabIndex = 13
        '
        'lblRecNo
        '
        Me.lblRecNo.Location = New System.Drawing.Point(192, 208)
        Me.lblRecNo.Name = "lblRecNo"
        Me.lblRecNo.TabIndex = 15
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(56, 208)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(64, 23)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "Record No.:"
        '
        'lblClientRec
        '
        Me.lblClientRec.Location = New System.Drawing.Point(192, 232)
        Me.lblClientRec.Name = "lblClientRec"
        Me.lblClientRec.TabIndex = 17
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(56, 232)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(120, 23)
        Me.Label8.TabIndex = 16
        Me.Label8.Text = "Total Client Records:"
        '
        'lblSkipped
        '
        Me.lblSkipped.Location = New System.Drawing.Point(192, 256)
        Me.lblSkipped.Name = "lblSkipped"
        Me.lblSkipped.TabIndex = 19
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(56, 256)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(128, 23)
        Me.Label9.TabIndex = 18
        Me.Label9.Text = "Clients with no Interest:"
        '
        'lblTimePassed
        '
        Me.lblTimePassed.Location = New System.Drawing.Point(432, 256)
        Me.lblTimePassed.Name = "lblTimePassed"
        Me.lblTimePassed.Size = New System.Drawing.Size(136, 23)
        Me.lblTimePassed.TabIndex = 21
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(336, 256)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(80, 23)
        Me.Label10.TabIndex = 20
        Me.Label10.Text = "Time Elapsed:"
        '
        'Timer1
        '
        '
        'SARSIT3Form
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(648, 282)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblTimePassed, Me.Label10, Me.Label9, Me.lblSkipped, Me.lblClientRec, Me.Label8, Me.lblRecNo, Me.Label7, Me.lblFileNo, Me.Label5, Me.Period, Me.btnExit, Me.btnExecute, Me.Label4, Me.txtPath, Me.Label3, Me.Label2, Me.Label1, Me.txtReference, Me.txtTest})
        Me.Name = "SARSIT3Form"
        Me.Text = "SARS Third Party Data Submission"
        Me.ResumeLayout(False)

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

    Public ReadOnly Property PeriodStart() As String
        Get
            Return MyDateString(DateAdd(DateInterval.Day, 1, DateAdd(DateInterval.Year, -1, CDate("#" & Period.Value.Month & "/" & Period.Value.Day & "/" & Period.Value.Year & "#"))))
        End Get
    End Property

    Public ReadOnly Property TestFile() As String
        Get
            Return txtTest.Text
        End Get
    End Property

    Public ReadOnly Property Reference() As String
        Get
            Return txtReference.Text
        End Get
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

    Public WriteOnly Property TimeElfapsed() As String
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

        If MessageBox.Show("Any existing files in the selected location will be overwritten.  Continue?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Timer1.Enabled = True
            CreateExportFile()
            Timer1.Enabled = False
            If Abort Then
                MessageBox.Show("File export aborted by user", "Done", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                MessageBox.Show("File created successfully", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If
    End Sub

    Private Sub SARSIT3Form_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        txtTest.Text = TestValue
        txtReference.Text = "GBS" & Period.Value.Year & FixedLengthString(Period.Value.Month, 2, "Right", "0") & FixedLengthString(Period.Value.Day, 2, "Right", "0")
        txtPath.Text = "\\DomainCtrl\GBSDocs$\DJBSmith\SARS IT3(b) Interface\"
    End Sub

    Private Sub Period_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Period.ValueChanged

        txtReference.Text = "GBS" & Period.Value.Year & FixedLengthString(Period.Value.Month, 2, "Right", "0") & FixedLengthString(Period.Value.Day, 2, "Right", "0")
    End Sub

    Private Sub txtTest_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTest.TextChanged

        If Not (txtTest.Text.ToUpper = "Y" Or txtTest.Text.ToUpper = "N" Or txtTest.Text = "") Then
            txtTest.Text = TestValue
        Else
            TestValue = txtTest.Text.ToUpper
            txtTest.Text = TestValue
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Dim SomeTime As Date
        Dim SecondsPassed As Integer
        Dim myHours As Integer
        Dim myMinutes As Integer

        SomeTime = Now
        SecondsPassed = DateDiff(DateInterval.Second, StartTime, SomeTime)
        myHours = SecondsPassed \ 3600
        SecondsPassed = SecondsPassed - (myhours * 3600)
        myMinutes = SecondsPassed \ 60
        SecondsPassed = SecondsPassed - (myMinutes * 60)
        lblTimePassed.Text = Format(myHours, "00") & ":" & Format(myMinutes, "00") & ":" & Format(SecondsPassed, "00")
    End Sub

End Class
