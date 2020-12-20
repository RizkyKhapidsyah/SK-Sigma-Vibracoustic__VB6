VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmLeaves 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Leaves Record"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9195
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   4200
      Top             =   3960
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=PMS"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "PMS"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT EMPCODE, YR, MN, SUM(LEAVDAY) FROM TRANLEAV where empcode = '"" & Trim(txtCode) & ""' GROUP BY empcode, YR, MN"
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   2160
      TabIndex        =   26
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   255
      Left            =   480
      TabIndex        =   29
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   255
      Left            =   1320
      TabIndex        =   28
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   255
      Left            =   3000
      TabIndex        =   27
      Top             =   3360
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   7680
      Top             =   3360
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=PMS"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "PMS"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Query1"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Bindings        =   "frmLeaves.frx":0000
      Height          =   3135
      Left            =   4320
      TabIndex        =   25
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5530
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4320
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=PMS"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "PMS"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmLeaves.frx":0015
      Height          =   1575
      Left            =   0
      TabIndex        =   24
      Top             =   3720
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   2778
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4215
      Begin MSComCtl2.DTPicker txtTo 
         Height          =   315
         Left            =   2640
         TabIndex        =   23
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   36370
      End
      Begin MSComCtl2.DTPicker txtFrom 
         Height          =   315
         Left            =   840
         TabIndex        =   22
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   36370
      End
      Begin VB.ComboBox cmbDay 
         Height          =   315
         Left            =   2640
         TabIndex        =   15
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtCode 
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbLType 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtOut 
         Height          =   315
         Left            =   2640
         TabIndex        =   3
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtIn 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Medical"
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   21
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Earned"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   20
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Casual"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   19
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lblCas 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   18
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label lblCas 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   17
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label lblCas 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   16
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Day"
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Code"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "L Type"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Out"
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Time In"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "To"
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Dt From"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "&Graph"
      Height          =   255
      Left            =   5880
      TabIndex        =   5
      Top             =   3360
      Width           =   1575
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Bindings        =   "frmLeaves.frx":002A
      Height          =   3135
      Left            =   4320
      OleObjectBlob   =   "frmLeaves.frx":003F
      TabIndex        =   30
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmLeaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbpms As Database
Dim rstTL As Recordset
Dim rstMP As Recordset

Private Sub cmbDay_Change()
Dim strMsg As String
Dim i As Integer
strMsg = ""
If cmbLType = "E" Or cmbLType = "C" Then
If cmbLType = "E" Then
    i = 1
Else
    i = 0
End If
If txtFrom = txtTo Then
Select Case cmbDay
Case "One"
If (Val(lblCas(i).Caption) - 1) < 0 Then
    strMsg = "0"
End If
Case "Half Day"
If (Val(lblCas(i).Caption) - 0.5) < 0 Then
    strMsg = "1"
End If
Case "One Third"
If (Val(lblCas(i).Caption) - 0.33) < 0 Then
    strMsg = "2"
End If
End Select
Else
If (Val(lblCas(i).Caption) - DateDiff("d", txtTo, txtFrom)) < 0 Then
    strMsg = "0"
End If
End If
If strMsg <> "" Then
    cmbLType.SetFocus
    MsgBox "No More Leaves Remaining. Sorry for Casual / Earned Leave "
End If
End If
End Sub


Private Sub cmdAdd_Click()
On Error GoTo ErrorHandler
Dim intTmp As Single
If txtCode = "" Then
    txtCode.SetFocus
    Exit Sub
End If
If txtFrom = "" Then
    MsgBox "Enter Date From and Date To"
    txtFrom.SetFocus
    Exit Sub
End If
If cmbLType.Text = "" Then
    cmbLType.SetFocus
    Exit Sub
End If
If txtTo = "" Then
    txtTo = txtFrom
End If
If txtIn = "" Then
    txtIn = "09:00:00"
End If
If txtOut = "" Then
    txtOut = "18:30:00"
End If

Select Case cmbDay
Case "One"
If (Val(lblCas(1).Caption) - 1) < 0 Then
    cmbDay.SetFocus
    Exit Sub
End If
Case "Half Day"
If (Val(lblCas(1).Caption) - 0.5) < 0 Then
    cmbDay.SetFocus
    Exit Sub
End If
Case "One Third"
If (Val(lblCas(1).Caption) - 0.33) < 0 Then
    cmbDay.SetFocus
    Exit Sub
End If
End Select
Dim mDt1 As Date
Dim mDt2 As Date
mDt1 = CDate(txtFrom)
mDt2 = CDate(txtTo)

Dim dbpms As Database
Dim rstTL As Recordset
Set dbpms = OpenDatabase(strPath & "dbpms.mdb")
Set rstTL = dbpms.OpenRecordset("tranleav", dbOpenDynaset)

With rstTL

    While mDt1 <= mDt2
        .AddNew
        !EMPCODE = txtCode
        !YR = Year(mDt1)
        !MN = Month(mDt1)
        !DY = Day(mDt1)
        !DOL = mDt1
        !LEAVTYPE = Left(cmbLType, 1)
        !timein = txtIn
        !timeout = txtOut
        Select Case cmbDay.Text
        Case "One Day":
            intTmp = 1
        Case "Half Day":
            intTmp = 0.5
        Case "One Third":
            intTmp = 0.33
        End Select
        !LEAVDAY = intTmp
        .Update
        mDt1 = DateAdd("d", 1, mDt1)
    Wend
    txtCode = ""
    txtName = ""
    txtFrom = Date
    txtTo = Date
    txtIn = "09:00:00"
    txtOut = "18:30:00"
    cmbLType = ""
    cmbDay = "One Day"
End With
Adodc2.RecordSource = "Select dol, empcode, leavtype, leavday from tranleav order by dol"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
    Adodc2.Recordset.MoveLast
End If
Exit Sub
ErrorHandler:
MsgBox "[" & Err.Number & "]  " & Err.Description
Unload Me
End Sub

Private Sub cmdCancel_Click()
txtCode = ""
txtName = ""
txtFrom = Format$(CStr(Now), "mm-dd-yyyy")
txtTo = Format$(CStr(Now), "mm-dd-yyyy")
cmbLType = ""
txtIn = "09:00:00"
txtOut = "18:00:00"
cmbDay = "One Day"

End Sub

Private Sub cmdDelete_Click()
If txtCode = "" Then
    txtCode.SetFocus
    Exit Sub
End If
If txtFrom = "" Then
    MsgBox "Enter Date From and Date To"
    txtFrom.SetFocus
    Exit Sub
End If
If cmbLType.Text = "" Then
    cmbLType.SetFocus
    Exit Sub
End If
If txtTo = "" Then
    txtTo = txtFrom
End If
If txtIn = "" Then
    txtIn = "09:00:00"
End If
If txtOut = "" Then
    txtOut = "18:30:00"
End If

Dim dbpms As Database
Dim rstTL As Recordset
Dim mdt As Date
mdt = CDate(txtFrom)
Set dbpms = OpenDatabase(strPath & "dbpms.mdb")
Set rstTL = dbpms.OpenRecordset("tranleav", dbOpenDynaset)
With rstTL
    .FindFirst "empcode = '" & txtCode & "' and yr = " & Year(mdt) & " and mn = " & Month(mdt) & " and dy = " & Day(mdt)
    If Not .EOF Then
        If MsgBox("Do you want to Delete this Record ", vbYesNo) = vbYes Then
            .Delete
        End If
    End If
End With
Adodc2.Refresh
End Sub

Private Sub cmdDetails_Click()
Dim ir As Integer
    If cmdDetails.Caption = "&Graph" Then
        cmdDetails.Caption = "D&etails"
        MSHFlexGrid2.Visible = False
        MSChart1.Top = MSHFlexGrid2.Top
        MSChart1.Left = MSHFlexGrid2.Left
        MSChart1.Width = MSHFlexGrid2.Width
        MSChart1.Height = MSHFlexGrid2.Height
        MSChart1.Visible = True
        Adodc3.CommandType = adCmdText
        Adodc3.RecordSource = "SELECT EMPCODE, YR, MN, SUM(LEAVDAY) FROM TRANLEAV where empcode = '" & Trim(txtCode) & "' GROUP BY empcode, YR, MN"
        Adodc3.Refresh
        MSChart1.chartType = VtChChartType2dBar
        With Adodc3.Recordset
            If .RecordCount > 0 Then
                MSChart1.RowCount = 1
                MSChart1.ColumnCount = 12
                MSChart1.RowLabel = "Months"
                ir = 1
                For ir = 1 To 12
                    MSChart1.Row = 1
                    MSChart1.Column = ir
                    MSChart1.Data = 0
                    Select Case ir
                    Case 1:
                    MSChart1.ColumnLabel = "Jan"
                    Case 2:
                    MSChart1.ColumnLabel = "Feb"
                    Case 3:
                    MSChart1.ColumnLabel = "Mar"
                    Case 4:
                    MSChart1.ColumnLabel = "Apr"
                    Case 5:
                    MSChart1.ColumnLabel = "May"
                    Case 6:
                    MSChart1.ColumnLabel = "Jun"
                    Case 7:
                    MSChart1.ColumnLabel = "Jul"
                    Case 8:
                    MSChart1.ColumnLabel = "Aug"
                    Case 9:
                    MSChart1.ColumnLabel = "Sep"
                    Case 10:
                    MSChart1.ColumnLabel = "Oct"
                    Case 11:
                    MSChart1.ColumnLabel = "Nov"
                    Case 12:
                    MSChart1.ColumnLabel = "Dec"
                    End Select
                Next ir
                ir = 1
                Do While Not .EOF
                    MSChart1.Row = 1
                    MSChart1.Column = .Fields(2)
                    MSChart1.Data = .Fields(3)
                    ir = ir + 1
                    .MoveNext
                    If .EOF Or .BOF Then
                        Exit Do
                    End If
                Loop
                MSChart1.ShowLegend = True
            End If
        End With
    Else
        cmdDetails.Caption = "&Graph"
        MSChart1.Visible = False
        MSHFlexGrid2.Visible = True
    End If
End Sub

Private Sub cmdExit_Click()
    Unload frmLeaves
    frmMain.Show
    Close
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler
Dim strSql As String
Set dbpms = OpenDatabase(strPath & "dbpms.mdb")
Set rstTL = dbpms.OpenRecordset("tranleav", dbOpenDynaset)
Set rstMP = dbpms.OpenRecordset("Mastpay", dbOpenDynaset)

MSHFlexGrid1.ColWidth(1) = 2000
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = mnuViewOrder
Adodc1.Refresh
Adodc2.CommandType = adCmdText
Adodc2.RecordSource = "Select dol, empcode, leavtype, leavday from tranleav order by dol"
Adodc2.Refresh

txtFrom.Value = DateAdd("m", -1, Now)
txtTo.Value = DateAdd("m", -1, Now)
txtIn.Text = "09:00:00"
txtOut.Text = "18:30:00"
cmbLType.AddItem ("Absent")
cmbLType.AddItem ("Casual")
cmbLType.AddItem ("Earned")
cmbLType.AddItem ("Field Duty")
cmbLType.AddItem ("L.W.Pay")
cmbLType.AddItem ("Holiday")
cmbLType.AddItem ("Present")
cmbDay.AddItem ("One Day")
cmbDay.AddItem ("Half Day")
cmbDay.AddItem ("One Third")
cmbDay.Text = "One Day"

Exit Sub
ErrorHandler:
MsgBox "[" & Err.Number & "]  " & Err.Description
Unload Me

End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then

End If
End Sub



Private Sub MSHFlexGrid1_DblClick()
MSHFlexGrid1.Col = 0
txtCode = MSHFlexGrid1.Text
MSHFlexGrid1.Col = 1
txtName = MSHFlexGrid1.Text
txtCode.SetFocus
Call txtCode_KeyPress(13)
End Sub

Private Sub MSHFlexGrid2_DblClick()
txtCode = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Row, 1)
txtFrom = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Row, 0)
txtTo = txtFrom
cmbLType = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Row, 2)
cmbDay = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Row, 3)

Select Case cmbLType
Case "A"
cmbLType = "Absent"
Case "F"
cmbLType = "Field Duty"
Case "C"
cmbLType = "Casual"
Case "E"
cmbLType = "Earned"
Case "L"
cmbLType = "L.W.Pay"
End Select

Select Case CDbl(cmbDay)
Case 1
cmbDay = "One"
Case 0.5
cmbDay = "Half Day"
Case 0.33
cmbDay = "One Third"
End Select
End Sub


Private Sub txtCode_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandler
Dim strString As String

If KeyAscii = 13 Then
    Set dbpms = OpenDatabase(strPath & "dbpms.mdb")
    Set rstMP = dbpms.OpenRecordset("select * from Mastpay", dbOpenDynaset)
    Set rstTL = dbpms.OpenRecordset("select * from tranleav", dbOpenDynaset)

    Dim strSql As String
    Dim sngEL, sngCL, sngSL As Integer
    Dim dtDOJ, dtCur, dtDOJ1 As Date
    Dim intTotDay As Integer
    lblCas(0).Caption = ""
    lblCas(1).Caption = ""
    
    If txtCode <> "" Then
        With rstMP
        rstMP.FindFirst "Empcode = '" & txtCode & "'"
        If Not .NoMatch Then
            txtName = !EMPNAME
            sngCL = !empcltot
            sngEL = !empeltot
            sngSL = 0
            dtDOJ = !EMPDOJ
            dtCur = Now
            dtDOJ1 = dtDOJ
            
            intTotDay = 0
            ' Calculating Casual Leaves
            If Year(dtDOJ) = Year(dtCur) Then
                dtDOJ1 = dtDOJ
                While dtDOJ1 <= CDate("12/31/" & Year(dtCur))
                    intTotDay = intTotDay + 1
                    dtDOJ1 = DateAdd("d", 1, dtDOJ1)
                Wend
                sngCL = (intTotDay / 365) * 7
            Else
                dtDOJ1 = CDate("01/01/" & Year(dtDOJ1))
                sngCL = 7
            End If
            
            
            intTotDay = 0
            ' Calculating Earned Leaves
            While dtDOJ <= dtCur
                If Weekday(dtDOJ) <> vbSunday Then
                    intTotDay = intTotDay + 1
                End If
                dtDOJ = DateAdd("d", 1, dtDOJ)
            Wend
            With rstTL
                .FindFirst "Empcode = '" & txtCode & "'"
                While Not .NoMatch
                    If Weekday(.Fields("dol")) <> vbSunday Then
                        Select Case .Fields("leavtype")
                        Case "E"
                            intTotDay = intTotDay - .Fields("leavday")
                        Case "C"
                            sngCL = sngCL - .Fields("Leavday")
                        End Select
                    End If
                    .FindNext "Empcode = '" & txtCode & "'"
                Wend
            End With
            sngEL = (intTotDay / 20)
            If (sngEL - Int(sngEL)) > 0.66 Then
                sngEL = Int(sngEL) + 1
            Else
                If (sngEL - Int(sngEL)) > 0.33 Then
                    sngEL = Int(sngEL) + 0.33
                Else
                    sngEL = Int(sngEL)
                End If
            End If
            
            If (sngCL - Int(sngCL)) >= 0.5 Then
                sngCL = Int(sngCL) + 1
            Else
                sngCL = Int(sngCL)
            End If
            
            lblCas(0).Caption = Format$(sngCL, "#.00")
            lblCas(1).Caption = Format$(sngEL, "#.00")
            strSql = "Select dol, empcode, leavtype, leavday from tranleav where empcode = '" & txtCode & "' order by dol"
            Adodc2.CommandType = adCmdText
            Adodc2.RecordSource = strSql
            Adodc2.Refresh
        End If
    End With
    End If
Else
    strString = Chr$(KeyAscii)
    strString = UCase$(strString)
    KeyAscii = Asc(strString)
End If
Exit Sub
ErrorHandler:
MsgBox "[" & Err.Number & "]  " & Err.Description
Unload Me
End Sub
Private Sub txtFrom_Change()
txtTo.Value = txtFrom.Value
End Sub

Private Sub txtFrom_LostFocus()
txtTo.Value = txtFrom.Value
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
Dim strSql As String
If KeyAscii = 13 Then
    If txtName <> " " Then
    End If
End If
End Sub

Private Sub chkdata()
End Sub
