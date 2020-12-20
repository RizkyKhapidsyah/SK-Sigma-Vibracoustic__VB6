VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmAdvance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Advance Feeding"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   1500
      TabIndex        =   7
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtRemark 
      Height          =   615
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox txtAmount 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   765
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   24576001
      CurrentDate     =   36495
   End
   Begin VB.TextBox txtEmpName 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   435
      Width           =   2655
   End
   Begin VB.TextBox txtEmpCode 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2400
      Top             =   0
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
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
      Bindings        =   "frmAdvance.frx":0000
      Height          =   2175
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0)._NumMapCols=   8
      _Band(0)._MapCol(0)._Name=   "empcode"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "empname"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "empdob"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "empdoj"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "empesi"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(5)._Name=   "emppf"
      _Band(0)._MapCol(5)._RSIndex=   5
      _Band(0)._MapCol(6)._Name=   "desgdes"
      _Band(0)._MapCol(6)._RSIndex=   6
      _Band(0)._MapCol(7)._Name=   "deptdes"
      _Band(0)._MapCol(7)._RSIndex=   7
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSG1 
      Bindings        =   "frmAdvance.frx":0015
      Height          =   1935
      Left            =   3840
      TabIndex        =   11
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      AllowBigSelection=   0   'False
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   5
      _Band(0)._MapCol(0)._Name=   "EMPCODE"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(0)._Hidden=   -1  'True
      _Band(0)._MapCol(1)._Name=   "DATE"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "TYPE"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "AMOUNT"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(3)._Alignment=   7
      _Band(0)._MapCol(4)._Name=   "REMARK"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(4)._Hidden=   -1  'True
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2400
      Top             =   1080
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
      RecordSource    =   "select * from advtab"
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
   Begin MSChart20Lib.MSChart MSChart1 
      Bindings        =   "frmAdvance.frx":002A
      Height          =   3015
      Left            =   3840
      OleObjectBlob   =   "frmAdvance.frx":003F
      TabIndex        =   17
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Remarks"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Width           =   630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Amount"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   1485
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Type"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1140
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   810
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Emp.Name"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "EmpCode "
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   165
      Width           =   735
   End
End
Attribute VB_Name = "frmAdvance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Dim ii As Integer
For ii = 1 To MSG1.Rows - 1
    
Next

With Adodc1.Recordset
    


End With
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
cmbType.AddItem "ADV"
cmbType.AddItem "I.T"
cmbType.AddItem "OTH"
cmbType.Text = "ADV"
dtDate = Date
MSHFlexGrid2.ColWidth(1) = 2000
Adodc2.CommandType = adCmdText
Adodc2.RecordSource = mnuViewOrder
Adodc2.Refresh
End Sub

Private Sub MSHFlexGrid2_DblClick()
txtEmpCode = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Row, 0)
txtEmpName = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Row, 1)
RetDataAdv
End Sub
Private Sub txtAmount_KeyPress(KeyAscii As Integer)
Dim strStr As String
If KeyPress <> 8 Then
    If InStr(1, txtAmount, ".") > 0 Then
        strStr = "0123456789"
    Else
        strStr = "0123456789."
    End If
    If InStr(1, strStr, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub txtEmpCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    For ii = 1 To MSHFlexGrid2.Rows - 1
        If Trim(MSHFlexGrid2.TextMatrix(ii, 0)) = Trim(txtEmpCode) Then
            txtEmpName = MSHFlexGrid2.TextMatrix(ii, 1)
            Exit Sub
        End If
    Next
    If ii = MSHFlexGrid2.Rows Then
        txtEmpName = ""
    End If
    RetDataAdv
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub RetDataAdv()
If txtEmpCode <> "" And txtEmpName <> "" Then
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "SELECT * from ADVTAB where empcode = '" & Trim(txtEmpCode) & "'"
    Adodc1.Refresh
End If
End Sub

Private Sub txtEmpCode_LostFocus()
txtEmpCode_KeyPress (13)
End Sub
