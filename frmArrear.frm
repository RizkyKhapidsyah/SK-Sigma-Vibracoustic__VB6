VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmArrear 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Arrears Feeding form"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3480
      Top             =   0
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
      CommandType     =   8
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
   Begin VB.TextBox txtSPL 
      Height          =   285
      Left            =   3600
      TabIndex        =   23
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtLTA 
      Height          =   285
      Left            =   3600
      TabIndex        =   22
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtMED 
      Height          =   285
      Left            =   3600
      TabIndex        =   21
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtDA 
      Height          =   285
      Left            =   3600
      TabIndex        =   20
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtUNI 
      Height          =   285
      Left            =   1320
      TabIndex        =   19
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtCONV 
      Height          =   285
      Left            =   1320
      TabIndex        =   18
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtHRA 
      Height          =   285
      Left            =   1320
      TabIndex        =   17
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtBasic 
      Height          =   285
      Left            =   1320
      TabIndex        =   16
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ComboBox cmbMonth 
      Height          =   315
      Left            =   3840
      TabIndex        =   13
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox cmbYear 
      Height          =   315
      Left            =   1560
      TabIndex        =   12
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox txtArrAmt 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtEmpCode 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.ComboBox cmbName 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   24
      Top             =   1920
      Width           =   4815
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Spl. All."
         Height          =   195
         Index           =   8
         Left            =   2880
         TabIndex        =   33
         Top             =   1800
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Uni. All."
         Height          =   195
         Index           =   7
         Left            =   600
         TabIndex        =   32
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "L.T.A."
         Height          =   195
         Index           =   6
         Left            =   2970
         TabIndex        =   31
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Conv. All"
         Height          =   195
         Index           =   5
         Left            =   510
         TabIndex        =   30
         Top             =   1440
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Medical"
         Height          =   195
         Index           =   4
         Left            =   2850
         TabIndex        =   29
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "H.R.A."
         Height          =   195
         Index           =   3
         Left            =   660
         TabIndex        =   28
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "D.A."
         Height          =   195
         Index           =   2
         Left            =   3090
         TabIndex        =   27
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Basic"
         Height          =   195
         Index           =   1
         Left            =   750
         TabIndex        =   26
         Top             =   720
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Arrear Amt."
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   25
         Top             =   360
         Width           =   780
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Arrear for the Year"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Month"
      Height          =   195
      Index           =   1
      Left            =   3240
      TabIndex        =   15
      Top             =   900
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Year"
      Height          =   195
      Index           =   0
      Left            =   1080
      TabIndex        =   14
      Top             =   900
      Width           =   330
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Spl.All."
      Height          =   195
      Index           =   7
      Left            =   3000
      TabIndex        =   8
      Top             =   3720
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "L.T.A."
      Height          =   195
      Index           =   5
      Left            =   2880
      TabIndex        =   7
      Top             =   3360
      Width           =   435
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Medical"
      Height          =   195
      Index           =   3
      Left            =   2880
      TabIndex        =   6
      Top             =   3000
      Width           =   555
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "D.A."
      Height          =   195
      Index           =   1
      Left            =   3000
      TabIndex        =   5
      Top             =   2640
      Width           =   315
   End
   Begin VB.Label Label2 
      Caption         =   "Employee Name"
      Height          =   255
      Left            =   315
      TabIndex        =   1
      Top             =   510
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Employee Code "
      Height          =   195
      Left            =   375
      TabIndex        =   0
      Top             =   165
      Width           =   1155
   End
End
Attribute VB_Name = "frmArrear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbMonth_Change()
GetData
End Sub

Private Sub cmbName_CLICK()
Dim strName As String
Dim ii As Integer
ii = InStr(1, cmbName, "{.")
If ii > 0 Then
    strName = Mid$(cmbName, 1, ii - 1)
End If
If strName <> "" Then
    txtEmpCode = ""
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "SELECT EMPCODE, EMPNAME, EMPDOJ FROM MASTPAY WHERE EMPNAME = '" & Trim(strName) & "'"
    Adodc1.Refresh
    With Adodc1.Recordset
        If .RecordCount > 0 Then
            txtEmpCode.Text = !EMPCODE
        End If
    End With
End If
GetData
End Sub
Private Sub GetData()
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "SELECT * FROM ARRTAB WHERE EMPCODE = '" & Trim(txtEmpCode) & "' AND YR = " & Val(cmbYear) & " AND MN = " & Val(cmbMonth)
Adodc1.Refresh
With Adodc1.Recordset
    If .RecordCount > 0 Then
        txtArrAmt = !ARRAMT
        txtBasic = !ARRBASIC
        txtDA = !ARRDA
        txtHRA = !ARRHRA
        txtMED = !ARRMED
        txtCONV = !ARRCA
        txtLTA = !ARRLTA
        txtUNI = !arruni
        txtSPL = !ARRSPL
    End If
End With
End Sub

Private Sub cmbYear_Change()
GetData
End Sub

Private Sub cmdCancel_Click()
txtEmpCode = ""
cmbName = ""
cmbYear = Year(Date)
cmbMonth = Month(Date)
txtArrAmt = ""
txtBasic = ""
txtDA = ""
txtHRA = ""
txtCONV = ""
txtMED = ""
txtLTA = ""
txtUNI = ""
txtSPL = ""
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdUpdate_Click()
If txtEmpCode = "" Then
    txtEmpCode.SetFocus
    Exit Sub
End If
If cmbYear = "" Then
    cmbYear.SetFocus
    Exit Sub
End If
If cmbMonth = "" Then
    cmbMonth.SetFocus
    Exit Sub
End If

Dim xBool As Boolean
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "SELECT empcode, empname FROM MASTPAY WHERE EMPCODE = '" & Trim(txtEmpCode) & "'"
Adodc1.Refresh
With Adodc1.Recordset
    If .RecordCount > 0 Then
        xBool = True
    End If
End With
If xBool = True Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "SELECT * FROM ARRTAB WHERE EMPCODE = '" & Trim(txtEmpCode) & "' AND YR = " & Val(cmbYear) & " AND MN = " & Val(cmbMonth)
    Adodc1.Refresh
    With Adodc1.Recordset
        If .RecordCount > 0 Then
            !ARRAMT = Val(txtArrAmt)
            !ARRBASIC = Val(txtBasic)
            !ARRDA = Val(txtDA)
            !ARRHRA = Val(txtHRA)
            !ARRMED = Val(txtMED)
            !ARRCA = Val(txtCONV)
            !ARRLTA = Val(txtLTA)
            !ARRUA = Val(txtUNI)
            !ARRSPL = Val(txtSPL)
        Else
        End If
    End With
End If
End Sub

Private Sub Form_Load()
Dim ii As Integer
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "SELECT EMPCODE, EMPNAME, EMPDOJ FROM MASTPAY WHERE empDOR <> #01/01/1999# ORDER BY EMPNAME"
Adodc1.Refresh
With Adodc1.Recordset
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            cmbName.AddItem Trim(!EMPNAME) & "{." & Format$(!EMPDOJ, "DD-MM-YYYY") & ".}"
            .MoveNext
        Loop
    End If
End With
For ii = 1999 To 2008
    cmbYear.AddItem ii
Next
For ii = 1 To 12
    cmbMonth.AddItem ii
Next
cmbYear.Text = Year(Date)
cmbMonth.Text = Month(Date)
End Sub


Private Sub txtEmpCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtEmpCode_LostFocus
End If
End Sub

Private Sub txtEmpCode_LostFocus()
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "SELECT EMPCODE, EMPNAME, EMPDOJ, EMPDOR FROM MASTPAY WHERE EMPCODE = '" & Trim(txtEmpCode) & "'"
Adodc1.Refresh
With Adodc1.Recordset
    If .RecordCount > 0 Then
        If !EMPDOR <> #1/1/1900# Then
            MsgBox "Employee Resigned. Unable to select"
            txtEmpCode = ""
            txtEmpCode.SetFocus
            Exit Sub
        Else
            cmbName = !EMPNAME & "{." & Format$(!EMPDOJ, "DD-MM-YYYY") & ".}"
        End If
    End If
End With
GetData
End Sub
