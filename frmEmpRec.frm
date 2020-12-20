VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEmpRec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Record"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\PMS\DBPMS\DBPMS.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "MASTPAY"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Modify"
      Height          =   375
      Left            =   1095
      TabIndex        =   4
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3285
      TabIndex        =   3
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2190
      TabIndex        =   2
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4320
      Width           =   1095
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmEmpRec.frx":0000
      Height          =   4215
      Left            =   0
      OleObjectBlob   =   "frmEmpRec.frx":0010
      TabIndex        =   0
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "frmEmpRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    strVal = "A"
    Load frmAddMod
    frmAddMod.Show
    Unload frmEmpRec
End Sub

Private Sub Command3_Click()
    Unload frmEmpRec
    frmMain.Enabled = True
    frmMain.Show
    Exit Sub
End Sub

Private Sub Command4_Click()
    strVal = "M"
    Load frmAddMod
    frmAddMod.Show
    Unload frmEmpRec
End Sub

Private Sub DBGrid1_SelChange(Cancel As Integer)
    If Data1.Recordset.RecordCount > 0 Then
    strVal1 = Data1.Recordset.Fields(0)
    End If
End Sub

Private Sub Form_Activate()
Form_Load
End Sub

Private Sub Form_Load()
    Dim strSql As String
    strSql = "select mastpay.empcode, mastpay.empname, mastpay.empdoj, mastpay.empesi, mastpay.emppf, " & _
    "mastdesg.desgdes, mastdept.deptdes from mastpay, mastdesg, mastdept " & _
    "where mastpay.empdesg = mastdesg.empdesg and mastpay.empdept = mastdept.empdept"
    
'    strSql = "Select * from mastpay "
    Data1.RecordSource = strSql
    Data1.Refresh
    If Data1.Recordset.RecordCount > 0 Then
    strVal1 = Data1.Recordset.Fields(0)
    Else
    strVal1 = "NO RECORD"
    End If
End Sub
