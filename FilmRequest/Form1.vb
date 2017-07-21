Imports System.Data.OleDb
Imports System.Text.RegularExpressions
Imports System.Configuration
Imports System.IO

Public Class Form1
    Inherits System.Windows.Forms.Form

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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents lblWoNo As System.Windows.Forms.Label
    Friend WithEvents txtWoNo As System.Windows.Forms.TextBox
    Friend WithEvents lblLotNo As System.Windows.Forms.Label
    Friend WithEvents txtLotNo As System.Windows.Forms.TextBox
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents lblScaleX As System.Windows.Forms.Label
    Friend WithEvents lblScaleY As System.Windows.Forms.Label
    Friend WithEvents txtScaleX As System.Windows.Forms.TextBox
    Friend WithEvents txtScaleY As System.Windows.Forms.TextBox
    Friend WithEvents lblDispToolNo As System.Windows.Forms.Label
    Friend WithEvents lblXDim As System.Windows.Forms.Label
    Friend WithEvents lblYDim As System.Windows.Forms.Label
    Friend WithEvents txtXDim As System.Windows.Forms.TextBox
    Friend WithEvents txtYDim As System.Windows.Forms.TextBox
    Friend WithEvents lblOpInit As System.Windows.Forms.Label
    Friend WithEvents txtOpInit As System.Windows.Forms.TextBox
    Friend WithEvents lblDispCustName As System.Windows.Forms.Label
    Friend WithEvents lblDispPartNo As System.Windows.Forms.Label
    Friend WithEvents btnFindIt As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents lblComments As System.Windows.Forms.Label
    Friend WithEvents lblProMSError As System.Windows.Forms.Label
    Friend WithEvents txtComments As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblWoNo = New System.Windows.Forms.Label
        Me.txtWoNo = New System.Windows.Forms.TextBox
        Me.lblLotNo = New System.Windows.Forms.Label
        Me.txtLotNo = New System.Windows.Forms.TextBox
        Me.btnFindIt = New System.Windows.Forms.Button
        Me.DataGrid1 = New System.Windows.Forms.DataGrid
        Me.lblScaleX = New System.Windows.Forms.Label
        Me.lblScaleY = New System.Windows.Forms.Label
        Me.txtScaleX = New System.Windows.Forms.TextBox
        Me.txtScaleY = New System.Windows.Forms.TextBox
        Me.lblDispToolNo = New System.Windows.Forms.Label
        Me.lblXDim = New System.Windows.Forms.Label
        Me.lblYDim = New System.Windows.Forms.Label
        Me.txtXDim = New System.Windows.Forms.TextBox
        Me.txtYDim = New System.Windows.Forms.TextBox
        Me.lblOpInit = New System.Windows.Forms.Label
        Me.txtOpInit = New System.Windows.Forms.TextBox
        Me.lblDispCustName = New System.Windows.Forms.Label
        Me.lblDispPartNo = New System.Windows.Forms.Label
        Me.btnExit = New System.Windows.Forms.Button
        Me.lblComments = New System.Windows.Forms.Label
        Me.txtComments = New System.Windows.Forms.TextBox
        Me.lblProMSError = New System.Windows.Forms.Label
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblWoNo
        '
        Me.lblWoNo.Location = New System.Drawing.Point(48, 8)
        Me.lblWoNo.Name = "lblWoNo"
        Me.lblWoNo.Size = New System.Drawing.Size(64, 16)
        Me.lblWoNo.TabIndex = 0
        Me.lblWoNo.Text = "Work Order"
        Me.lblWoNo.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'txtWoNo
        '
        Me.txtWoNo.Location = New System.Drawing.Point(120, 8)
        Me.txtWoNo.Name = "txtWoNo"
        Me.txtWoNo.Size = New System.Drawing.Size(128, 20)
        Me.txtWoNo.TabIndex = 1
        '
        'lblLotNo
        '
        Me.lblLotNo.Location = New System.Drawing.Point(80, 32)
        Me.lblLotNo.Name = "lblLotNo"
        Me.lblLotNo.Size = New System.Drawing.Size(32, 16)
        Me.lblLotNo.TabIndex = 2
        Me.lblLotNo.Text = "Lot"
        Me.lblLotNo.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'txtLotNo
        '
        Me.txtLotNo.Location = New System.Drawing.Point(120, 32)
        Me.txtLotNo.Name = "txtLotNo"
        Me.txtLotNo.Size = New System.Drawing.Size(40, 20)
        Me.txtLotNo.TabIndex = 3
        '
        'btnFindIt
        '
        Me.btnFindIt.Location = New System.Drawing.Point(328, 8)
        Me.btnFindIt.Name = "btnFindIt"
        Me.btnFindIt.Size = New System.Drawing.Size(112, 56)
        Me.btnFindIt.TabIndex = 16
        Me.btnFindIt.Text = "Find It"
        '
        'DataGrid1
        '
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(280, 8)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.ReadOnly = True
        Me.DataGrid1.Size = New System.Drawing.Size(40, 40)
        Me.DataGrid1.TabIndex = 18
        Me.DataGrid1.Visible = False
        '
        'lblScaleX
        '
        Me.lblScaleX.Location = New System.Drawing.Point(48, 128)
        Me.lblScaleX.Name = "lblScaleX"
        Me.lblScaleX.Size = New System.Drawing.Size(60, 23)
        Me.lblScaleX.TabIndex = 10
        Me.lblScaleX.Text = "Scale X"
        Me.lblScaleX.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblScaleY
        '
        Me.lblScaleY.Location = New System.Drawing.Point(48, 152)
        Me.lblScaleY.Name = "lblScaleY"
        Me.lblScaleY.Size = New System.Drawing.Size(60, 23)
        Me.lblScaleY.TabIndex = 12
        Me.lblScaleY.Text = "Scale Y"
        Me.lblScaleY.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtScaleX
        '
        Me.txtScaleX.Location = New System.Drawing.Point(120, 128)
        Me.txtScaleX.Name = "txtScaleX"
        Me.txtScaleX.Size = New System.Drawing.Size(100, 20)
        Me.txtScaleX.TabIndex = 11
        '
        'txtScaleY
        '
        Me.txtScaleY.Location = New System.Drawing.Point(120, 152)
        Me.txtScaleY.Name = "txtScaleY"
        Me.txtScaleY.Size = New System.Drawing.Size(100, 20)
        Me.txtScaleY.TabIndex = 13
        '
        'lblDispToolNo
        '
        Me.lblDispToolNo.Location = New System.Drawing.Point(120, 248)
        Me.lblDispToolNo.Name = "lblDispToolNo"
        Me.lblDispToolNo.Size = New System.Drawing.Size(100, 20)
        Me.lblDispToolNo.TabIndex = 19
        '
        'lblXDim
        '
        Me.lblXDim.Location = New System.Drawing.Point(16, 80)
        Me.lblXDim.Name = "lblXDim"
        Me.lblXDim.Size = New System.Drawing.Size(100, 23)
        Me.lblXDim.TabIndex = 6
        Me.lblXDim.Text = "X Dimension"
        Me.lblXDim.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblYDim
        '
        Me.lblYDim.Location = New System.Drawing.Point(16, 104)
        Me.lblYDim.Name = "lblYDim"
        Me.lblYDim.Size = New System.Drawing.Size(100, 23)
        Me.lblYDim.TabIndex = 8
        Me.lblYDim.Text = "Y Dimension"
        Me.lblYDim.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtXDim
        '
        Me.txtXDim.Location = New System.Drawing.Point(120, 80)
        Me.txtXDim.Name = "txtXDim"
        Me.txtXDim.Size = New System.Drawing.Size(40, 20)
        Me.txtXDim.TabIndex = 7
        '
        'txtYDim
        '
        Me.txtYDim.Location = New System.Drawing.Point(120, 104)
        Me.txtYDim.Name = "txtYDim"
        Me.txtYDim.Size = New System.Drawing.Size(40, 20)
        Me.txtYDim.TabIndex = 9
        '
        'lblOpInit
        '
        Me.lblOpInit.Location = New System.Drawing.Point(16, 56)
        Me.lblOpInit.Name = "lblOpInit"
        Me.lblOpInit.Size = New System.Drawing.Size(100, 23)
        Me.lblOpInit.TabIndex = 4
        Me.lblOpInit.Text = "Operator Initials"
        Me.lblOpInit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtOpInit
        '
        Me.txtOpInit.Location = New System.Drawing.Point(120, 56)
        Me.txtOpInit.Name = "txtOpInit"
        Me.txtOpInit.Size = New System.Drawing.Size(40, 20)
        Me.txtOpInit.TabIndex = 5
        '
        'lblDispCustName
        '
        Me.lblDispCustName.Location = New System.Drawing.Point(120, 264)
        Me.lblDispCustName.Name = "lblDispCustName"
        Me.lblDispCustName.Size = New System.Drawing.Size(280, 20)
        Me.lblDispCustName.TabIndex = 20
        '
        'lblDispPartNo
        '
        Me.lblDispPartNo.Location = New System.Drawing.Point(120, 280)
        Me.lblDispPartNo.Name = "lblDispPartNo"
        Me.lblDispPartNo.Size = New System.Drawing.Size(150, 20)
        Me.lblDispPartNo.TabIndex = 21
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(328, 72)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(112, 56)
        Me.btnExit.TabIndex = 17
        Me.btnExit.Text = "Exit"
        '
        'lblComments
        '
        Me.lblComments.Location = New System.Drawing.Point(56, 176)
        Me.lblComments.Name = "lblComments"
        Me.lblComments.Size = New System.Drawing.Size(60, 23)
        Me.lblComments.TabIndex = 14
        Me.lblComments.Text = "Comments"
        Me.lblComments.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtComments
        '
        Me.txtComments.Location = New System.Drawing.Point(120, 176)
        Me.txtComments.Multiline = True
        Me.txtComments.Name = "txtComments"
        Me.txtComments.Size = New System.Drawing.Size(208, 56)
        Me.txtComments.TabIndex = 15
        '
        'lblProMSError
        '
        Me.lblProMSError.AutoSize = True
        Me.lblProMSError.Location = New System.Drawing.Point(67, 262)
        Me.lblProMSError.Name = "lblProMSError"
        Me.lblProMSError.Size = New System.Drawing.Size(0, 13)
        Me.lblProMSError.TabIndex = 23
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(451, 326)
        Me.Controls.Add(Me.lblProMSError)
        Me.Controls.Add(Me.txtComments)
        Me.Controls.Add(Me.lblComments)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.lblDispPartNo)
        Me.Controls.Add(Me.lblDispCustName)
        Me.Controls.Add(Me.txtOpInit)
        Me.Controls.Add(Me.lblOpInit)
        Me.Controls.Add(Me.txtYDim)
        Me.Controls.Add(Me.txtXDim)
        Me.Controls.Add(Me.lblYDim)
        Me.Controls.Add(Me.lblXDim)
        Me.Controls.Add(Me.lblDispToolNo)
        Me.Controls.Add(Me.txtScaleY)
        Me.Controls.Add(Me.txtScaleX)
        Me.Controls.Add(Me.lblScaleY)
        Me.Controls.Add(Me.lblScaleX)
        Me.Controls.Add(Me.DataGrid1)
        Me.Controls.Add(Me.btnFindIt)
        Me.Controls.Add(Me.txtLotNo)
        Me.Controls.Add(Me.lblLotNo)
        Me.Controls.Add(Me.txtWoNo)
        Me.Controls.Add(Me.lblWoNo)
        Me.Name = "Form1"
        Me.Text = "Film Request"
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

#Region " Variables "
    Dim sPromsLocation As String = My.Settings("ProMSLocation")
    Dim sSMTPServer As String = My.Settings("SMTPServer")
    Dim sSender As String = My.Settings("Sender")
    Dim sRecipients As String = My.Settings("Recipients")
    Dim kfile As Boolean = False

#End Region

    Public Sub btnFindInfo_Click( _
    ByVal sender As System.Object, _
    ByVal e As System.EventArgs) _
    Handles btnFindIt.Click

        Dim ToolNo, Cust, PartNo, Rev, tmpWoNo, WoNo, ScaleX, ScaleY, XDim, YDim, OpInit, Comments As String
        Dim DTime As String = DateAndTime.Now

        Cursor = Cursors.WaitCursor     'Change cursor to busy
        lblDispCustName.Text = "Searching for Customer Information"     'Show that I am doing something
        Application.DoEvents()      'Refresh Form

        'If ProMS maintenance then exit
        If Check_KickFile(kfile) = True Then
            Application.Exit()
            Exit Sub
        End If

        
        'TODO Add validation for work order number
        'TODO Add validation for the rest of the form

        'If barcode wand was used, split out the lot number.
        If txtWoNo.Text.Length = 15 Then
            tmpWoNo = LTrim(txtWoNo.Text)
            txtLotNo.Text = Microsoft.VisualBasic.Right(txtWoNo.Text, 3)
            txtOpInit.Focus()
        Else
            'If typed, move to get lot number
            tmpWoNo = LTrim(txtWoNo.Text)
            txtLotNo.Focus()
        End If

        'Find the data, or send the mail
        If btnFindIt.Text = "Find It" Then
            FindIt(tmpWoNo)
        Else
            ToolNo = RTrim(DataGrid1.Item(0, 3))
            Cust = RTrim(DataGrid1.Item(0, 0))
            PartNo = RTrim(DataGrid1.Item(0, 1))
            Rev = RTrim(DataGrid1.Item(0, 2))
            WoNo = Microsoft.VisualBasic.Left(tmpWoNo, 7) & "-" & txtLotNo.Text
            ScaleX = txtScaleX.Text
            ScaleY = txtScaleY.Text
            XDim = txtXDim.Text
            YDim = txtYDim.Text
            OpInit = txtOpInit.Text.ToUpper
            Comments = txtComments.Text.ToUpper
            MailHelper.SendMailMessage(My.Settings.SMTPServer, My.Settings.Sender, My.Settings.Recipients, "", "", ToolNo + " " + WoNo + " Scaled Film Request", "<!DOCTYPE html PUBLIC " & "-//W3C//DTD HTML 4.01 Transitional//EN" & _
            "><html><head><meta content=" & "text/html; charset=ISO-8859-1" & _
            " http-equiv=" & "content-type" & "><title>Scaled Film Request</title>" & _
            "</head><body>" & "<div style=" & "text-align: center;" & ">" & _
            "<div style=" & "text-align: left;" & "><big><big><big><span style=" & _
            "font-weight: bold;" & ">Scaled Film Request</span></big></big></big><br>" & _
            "<div style=" & "text-align: left;" & "><big><small><br></small></big>" & _
            "<table style=" & "text-align: left; width: 713px;" & "border=" & "2" & "cellpadding=" & "2" & " cellspacing=" & "2" & ">" & _
            "<tbody><tr><td>Tool#</td><td>" & ToolNo & "</td></tr>" & _
            "<tr><td>Customer:</td><td>" & Cust & "</td></tr>" & _
            "<tr><td>Part Number:</td><td>" & PartNo & "  " & Rev & "</td></tr>" & _
            "<tr><td>WO#</td><td>" & WoNo & "</td></tr>" & _
            "<tr><td>Date/Time Requested:</td><td>" & DTime & "</td></tr>" & _
            "<tr><td>Operator Initials</td><td>" & OpInit & "</td></tr>" & _
            "<tr><td>Scale " & XDim & " inch Dimension:</td><td>" & ScaleX & "</td></tr>" & _
            "<tr><td>Scale " & YDim & " inch Dimension:</td><td>" & ScaleY & "</td></tr>" & _
            "<tr><td>Comments:</td><td>" & Comments & "</td></tr>" & _
            "</tbody></table></div></div></body></html>")
            ResetForm()
        End If

        Cursor = Cursors.Default        'Change cursor to normal
        Application.DoEvents()          'Refresh form

    End Sub 'btnFindInfo_Click

    Private Sub ResetForm()
        btnFindIt.Text = "Find It"
        txtWoNo.Text = ""
        txtLotNo.Text = ""
        txtScaleX.Text = ""
        txtScaleY.Text = ""
        txtXDim.Text = ""
        txtYDim.Text = ""
        txtOpInit.Text = ""
        lblDispToolNo.Text = ""
        lblDispCustName.Text = ""
        lblDispPartNo.Text = ""
        txtComments.Text = ""
        txtWoNo.Focus()
    End Sub 'ResetForm

    Public Sub FindIt(ByVal tmpWoNo As String)
        Cursor.Current = Cursors.WaitCursor '# for hourglass

        Dim FindWoNo As String = Microsoft.VisualBasic.Left(tmpWoNo, 7)    'Search using the base wo number

        Dim blnInputOK As Boolean = CheckWO(FindWoNo) 'Check for valid input

        If Not blnInputOK Then  'Check result for valid input
            Beep()
            txtWoNo.Clear()
            txtWoNo.Focus()
            Exit Sub
        End If


        'Create oledb connection string
        Dim sFoxProConnString As String = _
                "Provider=vfpoledb.1;" & _
                "Data Source=" & sPromsLocation & ";" & _
                "Collating Sequence=machine"

        'Assign connection string
        Dim oFoxProConn As OleDbConnection = New OleDbConnection(sFoxProConnString)

        'Place to hold results
        Dim dtFoxProDT As New DataTable

        'Open connection
        oFoxProConn.Open()

        'SQL command to get ProMS data      'select the fields from tables joined on fjobnum looking for base WO#
        Dim sFoxProSql As String = _
        "SELECT PM_DUED.fcusname, " & _
        "PM_DUED.fpartnum, " & _
        "PM_DUED.fpartrev, " & _
        "PM_BOOK.ftoolnum " & _
        "FROM PM_DUED " & _
        "JOIN PM_BOOK " & _
        "ON PM_DUED.fjobnum = PM_BOOK.fjobnum " & _
        "WHERE PM_DUED.fjobnum = " & FindWoNo

        'Adapter to get data
        Dim daFoxProAdapt _
        As New OleDbDataAdapter _
        (sFoxProSql, oFoxProConn)

        'Command to execute
        Dim oleCmd As OleDbCommand = _
        New OleDbCommand(sFoxProSql, oFoxProConn)

        oleCmd.ExecuteNonQuery()    'Find matching record
        daFoxProAdapt.Fill(dtFoxProDT)  'Fill data table with matches
        DataGrid1.DataSource = dtFoxProDT   'Display data in data table
        oFoxProConn.Close()     'Close connection
        oFoxProConn.Dispose()   'Take out the trash
        'dtFoxProDT.PrimaryKey = New DataColumn() {dtFoxProDT.Columns(3)}

        'Did we return a record?
        If dtFoxProDT.Rows.Count = 0 Then
            Beep()
            ResetForm()
        Else
            lblDispToolNo.Text = RTrim(DataGrid1.Item(0, 3))
            lblDispCustName.Text = Microsoft.VisualBasic.Left(RTrim(DataGrid1.Item(0, 0)), 30)
            lblDispPartNo.Text = Microsoft.VisualBasic.Left(RTrim(DataGrid1.Item(0, 1)), 20) & "  " & RTrim(DataGrid1.Item(0, 2))
            btnFindIt.Text = "Send It"
        End If

        Cursor.Current = Cursors.Default    '# for normal


    End Sub 'FindIt

    Public Sub SendHTMLMail( _
    ByVal ToolNo As String, _
    ByVal Cust As String, _
    ByVal PartNo As String, _
    ByVal Rev As String, _
    ByVal WoNo As String, _
    ByVal ScaleX As String, _
    ByVal ScaleY As String, _
    ByVal XDim As String, _
    ByVal YDim As String, _
    ByVal OpInit As String, _
    ByVal Comments As String, ByVal sSender As String, ByVal sRecipients As String)

        Dim DTime As String = DateAndTime.Now


    End Sub 'SendHTMLMail

    Private Function CheckWO( _
    ByVal FindWoNo As String) As String

        Dim woReg As String = "^\d{7}"   'match the first 7 digits
        Dim woRegOptions As RegexOptions = ((RegexOptions.IgnorePatternWhitespace _
        Or RegexOptions.Multiline) _
        Or RegexOptions.IgnoreCase)
        Dim woRegex As New Regex(woReg)
        Return woRegex.IsMatch(FindWoNo)

    End Function 'CheckWO

    Private Sub btnExit_Click( _
    ByVal sender As System.Object, _
    ByVal e As System.EventArgs) _
    Handles btnExit.Click

        Application.Exit()

    End Sub 'btnExit_Click

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Check for proper network connection
        Dim sDued As String = sPromsLocation & "\PM_DUED.dbf"
        If Not File.Exists(sDued) Then
            MessageBox.Show("Cannot connect to ProMS" & "," & " Please login!", "Network Error!")
            Application.Exit()
        End If
        If Check_KickFile(kfile) = True Then
            MessageBox.Show("ProMS is unavailable due to maintenance" & "," & " Please try again later!", "ProMS maintenance!")
            Application.Exit()
        End If

    End Sub 'Form1_Load

    Private Function Check_KickFile(ByVal kfile As Boolean) As Boolean

        If File.Exists(My.Settings.KickFile) Then kfile = True
        Return (kfile)

    End Function 'Check_KickFile

End Class 'Form1
