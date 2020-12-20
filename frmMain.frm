VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Sigma Vibracoustic (Personnel Management System)"
   ClientHeight    =   3870
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6375
   ScaleHeight     =   3870
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2520
      Top             =   1560
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
      RecordSource    =   "MASTPAY"
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
      Bindings        =   "frmMain.frx":0000
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   9975
      _Version        =   393216
      Cols            =   28
      FixedCols       =   0
      AllowUserResizing=   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   28
      _Band(0)._NumMapCols=   28
      _Band(0)._MapCol(0)._Name=   "EMPCODE"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "EMPNAME"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "EMPSEX"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "EMPFATH"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "EMPDOB"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(5)._Name=   "EMPBLG"
      _Band(0)._MapCol(5)._RSIndex=   5
      _Band(0)._MapCol(6)._Name=   "EMPPA1"
      _Band(0)._MapCol(6)._RSIndex=   6
      _Band(0)._MapCol(7)._Name=   "EMPPA2"
      _Band(0)._MapCol(7)._RSIndex=   7
      _Band(0)._MapCol(8)._Name=   "EMPPATEL"
      _Band(0)._MapCol(8)._RSIndex=   8
      _Band(0)._MapCol(9)._Name=   "EMPCA1"
      _Band(0)._MapCol(9)._RSIndex=   9
      _Band(0)._MapCol(10)._Name=   "EMPCA2"
      _Band(0)._MapCol(10)._RSIndex=   10
      _Band(0)._MapCol(11)._Name=   "EMPCATEL"
      _Band(0)._MapCol(11)._RSIndex=   11
      _Band(0)._MapCol(12)._Name=   "EMPINF"
      _Band(0)._MapCol(12)._RSIndex=   12
      _Band(0)._MapCol(13)._Name=   "EMPNOM"
      _Band(0)._MapCol(13)._RSIndex=   13
      _Band(0)._MapCol(14)._Name=   "EMPNOMREL"
      _Band(0)._MapCol(14)._RSIndex=   14
      _Band(0)._MapCol(15)._Name=   "EMPDOJ"
      _Band(0)._MapCol(15)._RSIndex=   15
      _Band(0)._MapCol(16)._Name=   "EMPDESG"
      _Band(0)._MapCol(16)._RSIndex=   16
      _Band(0)._MapCol(17)._Name=   "EMPDEPT"
      _Band(0)._MapCol(17)._RSIndex=   17
      _Band(0)._MapCol(18)._Name=   "EMPCOMP"
      _Band(0)._MapCol(18)._RSIndex=   18
      _Band(0)._MapCol(19)._Name=   "EMPPF"
      _Band(0)._MapCol(19)._RSIndex=   19
      _Band(0)._MapCol(20)._Name=   "EMPESI"
      _Band(0)._MapCol(20)._RSIndex=   20
      _Band(0)._MapCol(21)._Name=   "EMPRES"
      _Band(0)._MapCol(21)._RSIndex=   21
      _Band(0)._MapCol(22)._Name=   "EMPDOR"
      _Band(0)._MapCol(22)._RSIndex=   22
      _Band(0)._MapCol(23)._Name=   "EMPCLTOT"
      _Band(0)._MapCol(23)._RSIndex=   23
      _Band(0)._MapCol(23)._Alignment=   7
      _Band(0)._MapCol(24)._Name=   "EMPELTOT"
      _Band(0)._MapCol(24)._RSIndex=   24
      _Band(0)._MapCol(24)._Alignment=   7
      _Band(0)._MapCol(25)._Name=   "EMPSLTOT"
      _Band(0)._MapCol(25)._RSIndex=   25
      _Band(0)._MapCol(25)._Alignment=   7
      _Band(0)._MapCol(26)._Name=   "EMPHLTOT"
      _Band(0)._MapCol(26)._RSIndex=   26
      _Band(0)._MapCol(26)._Alignment=   7
      _Band(0)._MapCol(27)._Name=   "BANKACC"
      _Band(0)._MapCol(27)._RSIndex=   27
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5160
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0015
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0073
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":00D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":012F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":018D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":01EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0249
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":02A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0305
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuSeperator0 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "&Save &As"
      End
      Begin VB.Menu mnuSeperator1 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up"
      End
      Begin VB.Menu mnuFilePrintArea 
         Caption         =   "Prin&t Area"
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Pre&view"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuSeperator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuSeperator5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find"
      End
      Begin VB.Menu mnuEditReplace 
         Caption         =   "R&eplace"
      End
      Begin VB.Menu mnuEditGoto 
         Caption         =   "&Goto"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      WindowList      =   -1  'True
      Begin VB.Menu mnuViewName 
         Caption         =   "&Name"
         Index           =   0
      End
      Begin VB.Menu mnuViewName 
         Caption         =   "&Code"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuViewName 
         Caption         =   "&E.S.I"
         Index           =   2
      End
      Begin VB.Menu mnuViewName 
         Caption         =   "&P.F"
         Index           =   3
      End
      Begin VB.Menu mnuViewName 
         Caption         =   "&DOJ "
         Index           =   4
      End
      Begin VB.Menu mnuViewName 
         Caption         =   "D&OB"
         Index           =   5
      End
   End
   Begin VB.Menu mnuOthers 
      Caption         =   "&Other"
      Begin VB.Menu mnuLeavRec 
         Caption         =   "&Leaves Record"
      End
      Begin VB.Menu mnuDesgRec 
         Caption         =   "&Desg. Record"
      End
      Begin VB.Menu mnuDeptRec 
         Caption         =   "D&ept Record"
      End
      Begin VB.Menu mnuArrRec 
         Caption         =   "&Arrear Record"
      End
      Begin VB.Menu mnuOthAdv 
         Caption         =   "&Advance/I.T./Other Ded"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Reports"
      Begin VB.Menu mnuSalmain 
         Caption         =   "Salary"
         Begin VB.Menu mnuSalCal 
            Caption         =   "&Sal. Calculation"
         End
         Begin VB.Menu mnuSalbank 
            Caption         =   "&Salary Print"
         End
      End
      Begin VB.Menu mnuRptPF 
         Caption         =   "&P.F.Register"
         Begin VB.Menu mnuRptPFMly 
            Caption         =   "&Monthly"
         End
         Begin VB.Menu mnuRptPFQly 
            Caption         =   "&Quaterly"
         End
         Begin VB.Menu mnuRptPFHly 
            Caption         =   "&Half Yearly"
         End
         Begin VB.Menu mnuRptPFYly 
            Caption         =   "&Yearly"
         End
      End
      Begin VB.Menu mnuRptESI 
         Caption         =   "&E.S.I."
         Begin VB.Menu mnuRptESIMly 
            Caption         =   "&Monthly"
         End
         Begin VB.Menu mnuRptESIQly 
            Caption         =   "&Quaterly"
         End
         Begin VB.Menu mnuRptESIHly 
            Caption         =   "&Half Yearly"
         End
         Begin VB.Menu mnuRptESIYly 
            Caption         =   "&Yearly"
         End
      End
      Begin VB.Menu mnuRPTLV 
         Caption         =   "&Leaves Report"
         Begin VB.Menu mnuRptLVDly 
            Caption         =   "&Daily All"
         End
         Begin VB.Menu mnuRPTLVDEPT 
            Caption         =   "D&epartment Wise"
         End
         Begin VB.Menu mnuRPTLVEMP 
            Caption         =   "&Employee Wise"
         End
      End
   End
   Begin VB.Menu mnuHoli 
      Caption         =   "&Holidays"
      Begin VB.Menu mnuHol 
         Caption         =   "&Holiday Add/Delete"
         Index           =   0
      End
   End
   Begin VB.Menu mnuTransfer 
      Caption         =   "&Transfer"
      Begin VB.Menu mnuHdOffTr 
         Caption         =   "&File To HeadOffice"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuWindowPath 
         Caption         =   "&Path"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
Adodc1.RecordSource = mnuViewOrder
Adodc1.Refresh
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler
    Me.mnuFileSave.Enabled = False
    Me.mnuFileSaveAs.Enabled = False
    MSHFlexGrid1.ColWidth(1) = 2000
    MSHFlexGrid1.ColWidth(6) = 2000
    MSHFlexGrid1.Visible = False
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = mnuViewOrder
    Adodc1.Refresh
Exit Sub
ErrorHandler:
MsgBox "[" & Err.Number & "]  " & Err.Description
Unload Me
End Sub

Private Sub Form_Resize()
Call frmHeightWidth
End Sub


Private Sub mnuArrRec_Click()
Load frmArrear
frmArrear.Show 1
End Sub

Private Sub mnuDeptRec_Click()
Load frmMASTDEPT
frmMASTDEPT.Show 1
End Sub

Private Sub mnuDesgRec_Click()
Load frmMASTDESG
frmMASTDESG.Show 1
End Sub

Private Sub mnuEditFind_Click()
Load frmFind
frmFind.Show 1
End Sub

Private Sub mnuFileClose_Click()
frmMain.Show
Unload frmAddMod
MSHFlexGrid1.Visible = False
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileNew_Click()
    strVal = "A"
    Load frmAddMod
    frmAddMod.Show 1
End Sub

Private Sub mnuFileOpen_Click()
Dim strVar As String
MSHFlexGrid1.Visible = True
Call frmHeightWidth
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = mnuViewOrder
Adodc1.Refresh
End Sub
Private Sub frmHeightWidth()
If Me.ScaleHeight >= 6550 Then
MSHFlexGrid1.Height = (Me.ScaleHeight - 1600)
Else
    If Me.ScaleHeight > 500 Then
        MSHFlexGrid1.Height = (Me.ScaleHeight - 500)
    Else
        MSHFlexGrid1.Visible = False
    End If
End If
If Me.ScaleWidth >= 9600 Then
    MSHFlexGrid1.Width = (Me.ScaleWidth - 200)
Else
    If MSHFlexGrid1.Width < 100 Then
        MSHFlexGrid1.Width = (Me.ScaleWidth - 100)
    Else
        MSHFlexGrid1.Visible = False
    End If
End If
End Sub


Private Sub mnuHol_Click(Index As Integer)
Select Case Index
Case 0
    Load frmHlAddDel
    frmHlAddDel.Show 1
End Select
End Sub

Private Sub mnuLeavRec_Click()
    Load frmLeaves
    frmLeaves.Show 1
End Sub

Private Sub mnuOthAdv_Click()
Load frmAdvance
frmAdvance.Show 1
End Sub

Private Sub mnuRptLVDly_Click()
Load frmPFESI
frmPFESI.Show 1
End Sub

Private Sub mnuRptPFMly_Click()
Load frmPFESI
frmPFESI.Show 1
End Sub

Private Sub mnuSalbank_Click()
strVal = "2"
    Load frmCalSal
    frmCalSal.Show 1
End Sub

Private Sub mnuSalCal_Click()
    strVal = "1"
    Load frmCalSal
    frmCalSal.Show 1
End Sub

Private Sub mnuSalReg_Click()
    strVal = "2"
    Load frmCalSal
    frmCalSal.Show 1
End Sub

Private Sub mnuViewCode_Click(Index As Integer)
Call mnuviewchecked(Index)
End Sub

Private Sub mnuViewDOB_Click(Index As Integer)
Call mnuviewchecked(Index)
End Sub

Private Sub mnuViewDOJ_Click(Index As Integer)
Call mnuviewchecked(Index)
End Sub

Private Sub mnuViewESI_Click(Index As Integer)
Call mnuviewchecked(Index)
End Sub

Private Sub mnuViewName_Click(Index As Integer)
Call mnuviewchecked(Index)
End Sub
Private Sub mnuviewchecked(mIndex As Integer)
Dim i As Integer
Dim strVar As String
For i = 0 To 5 Step 1
If mnuViewName(i).Checked Then
    mnuViewName(i).Checked = False
End If
Next
mnuViewName(mIndex).Checked = True
Adodc1.RecordSource = mnuViewOrder
Adodc1.Refresh
End Sub


Private Sub mnuViewPF_Click(Index As Integer)
Call mnuviewchecked(Index)
End Sub

Private Sub mnuWindowPath_Click()
Dim strPass As String
strPass = InputBox("Enter Password to Change the Windows Path ", "Password", , 90, 200)
If strPass = "mandhi" Then
    strPath = InputBox("Enter Path : ", , "c:\pms\dbpms\", 100, 200)
End If
End Sub

Private Sub MSHFlexGrid1_DblClick()
MSHFlexGrid1.Col = 0
strVal1 = MSHFlexGrid1.Text
strVal = "M"
Load frmAddMod
frmAddMod.Show 1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "New"
    Call mnuFileNew_Click
Case "Open"
    Call mnuFileOpen_Click
Case "Save"
Case "Print"
Case "Undo"
Case "Cut"
Case "Copy"
Case "Paste"
End Select
End Sub
