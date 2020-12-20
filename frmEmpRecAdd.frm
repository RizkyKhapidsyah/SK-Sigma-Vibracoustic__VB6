VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAddMod 
   Caption         =   "Add Employee Record"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "FoxPro 2.6;"
      DatabaseName    =   "C:\PMS\pmsdbf"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Mastsal"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "FoxPro 2.6;"
      DatabaseName    =   "C:\PMS\pmsdbf"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Mastcomp"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "FoxPro 2.6;"
      DatabaseName    =   "C:\PMS\pmsdbf"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Mastdesg"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "FoxPro 2.6;"
      DatabaseName    =   "C:\PMS\pmsdbf"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Mastpay"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "FoxPro 2.6;"
      DatabaseName    =   "C:\PMS\pmsdbf"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Mastdept"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   4920
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   8281
      _Version        =   327680
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Emp.Details"
      TabPicture(0)   =   "frmEmpRecAdd.frx":0000
      Tab(0).ControlCount=   20
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label10"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtCode"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtName"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtFathName"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtBLG"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtDOJ"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmbDesg"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmbComp"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmbSex"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtDOB"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Combo1"
      Tab(0).Control(19).Enabled=   0   'False
      TabCaption(1)   =   "Adv. Details"
      TabPicture(1)   =   "frmEmpRecAdd.frx":001C
      Tab(1).ControlCount=   14
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label21"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label22"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label23"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label24"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label25"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtCA1"
      Tab(1).Control(6).Enabled=   -1  'True
      Tab(1).Control(7)=   "txtCA2"
      Tab(1).Control(7).Enabled=   -1  'True
      Tab(1).Control(8)=   "txtCATel"
      Tab(1).Control(8).Enabled=   -1  'True
      Tab(1).Control(9)=   "txtPA1"
      Tab(1).Control(9).Enabled=   -1  'True
      Tab(1).Control(10)=   "txtPA2"
      Tab(1).Control(10).Enabled=   -1  'True
      Tab(1).Control(11)=   "txtPATel"
      Tab(1).Control(11).Enabled=   -1  'True
      Tab(1).Control(12)=   "txtNom"
      Tab(1).Control(12).Enabled=   -1  'True
      Tab(1).Control(13)=   "txtNomRel"
      Tab(1).Control(13).Enabled=   -1  'True
      TabCaption(2)   =   "Salary"
      TabPicture(2)   =   "frmEmpRecAdd.frx":0038
      Tab(2).ControlCount=   19
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label15"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label14"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label13"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label12"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label11"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label16"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label17"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label18"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label19"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "txtUALL"
      Tab(2).Control(9).Enabled=   -1  'True
      Tab(2).Control(10)=   "txtCALL"
      Tab(2).Control(10).Enabled=   -1  'True
      Tab(2).Control(11)=   "txtHRA"
      Tab(2).Control(11).Enabled=   -1  'True
      Tab(2).Control(12)=   "txtBasic"
      Tab(2).Control(12).Enabled=   -1  'True
      Tab(2).Control(13)=   "txtGross"
      Tab(2).Control(13).Enabled=   -1  'True
      Tab(2).Control(14)=   "txtDA"
      Tab(2).Control(14).Enabled=   -1  'True
      Tab(2).Control(15)=   "txtMedical"
      Tab(2).Control(15).Enabled=   -1  'True
      Tab(2).Control(16)=   "txtLTA"
      Tab(2).Control(16).Enabled=   -1  'True
      Tab(2).Control(17)=   "txtSPLALL"
      Tab(2).Control(17).Enabled=   -1  'True
      Tab(2).Control(18)=   "Text1"
      Tab(2).Control(18).Enabled=   -1  'True
      TabCaption(3)   =   "ESI && PF"
      TabPicture(3)   =   "frmEmpRecAdd.frx":0054
      Tab(3).ControlCount=   4
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label26"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label27"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "txtESICode"
      Tab(3).Control(2).Enabled=   -1  'True
      Tab(3).Control(3)=   "txtPFCode"
      Tab(3).Control(3).Enabled=   -1  'True
      TabCaption(4)   =   "Leaves"
      TabPicture(4)   =   "frmEmpRecAdd.frx":0070
      Tab(4).ControlCount=   4
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label28"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label29"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "txtCL"
      Tab(4).Control(2).Enabled=   -1  'True
      Tab(4).Control(3)=   "txtEL"
      Tab(4).Control(3).Enabled=   -1  'True
      Begin VB.TextBox Text1 
         DataField       =   "EMPCODE"
         DataSource      =   "Data5"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -71160
         TabIndex        =   64
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtEL 
         DataField       =   "EMPELTOT"
         DataSource      =   "Data5"
         Height          =   375
         Left            =   -73200
         TabIndex        =   63
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtCL 
         DataField       =   "EMPCLTOT"
         DataSource      =   "Data5"
         Height          =   375
         Left            =   -73200
         TabIndex        =   61
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtPFCode 
         DataField       =   "EMPPF"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   -72960
         TabIndex        =   59
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtESICode 
         DataField       =   "EMPESI"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   -72960
         TabIndex        =   57
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtNomRel 
         DataField       =   "EMPNOMREL"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   -69600
         TabIndex        =   55
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox txtNom 
         DataField       =   "EMPNOM"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   -72960
         TabIndex        =   53
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox txtPATel 
         DataField       =   "EMPPATEL"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   -72960
         TabIndex        =   51
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox txtPA2 
         DataField       =   "EMPPA2"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   -72960
         TabIndex        =   49
         Top             =   2160
         Width           =   3615
      End
      Begin VB.TextBox txtPA1 
         DataField       =   "EMPPA1"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   -72960
         TabIndex        =   48
         Top             =   1800
         Width           =   3615
      End
      Begin VB.TextBox txtCATel 
         DataField       =   "EMPCATEL"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   -72960
         TabIndex        =   46
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtCA2 
         DataField       =   "EMPCA2"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   -72960
         TabIndex        =   44
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtCA1 
         DataField       =   "EMPCA1"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   -72960
         TabIndex        =   43
         Top             =   600
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "EMPDEPT"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1320
         TabIndex        =   41
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox txtSPLALL 
         DataField       =   "EMPSPLALL"
         DataSource      =   "Data5"
         Height          =   375
         Left            =   -71160
         TabIndex        =   40
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtLTA 
         DataField       =   "EMPLTALL"
         DataSource      =   "Data5"
         Height          =   375
         Left            =   -71160
         TabIndex        =   39
         Top             =   2310
         Width           =   1335
      End
      Begin VB.TextBox txtMedical 
         DataField       =   "EMPMED"
         DataSource      =   "Data5"
         Height          =   375
         Left            =   -71160
         TabIndex        =   38
         Top             =   1740
         Width           =   1335
      End
      Begin VB.TextBox txtDA 
         DataField       =   "EMPDA"
         DataSource      =   "Data5"
         Height          =   375
         Left            =   -71160
         TabIndex        =   37
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtGross 
         Alignment       =   1  'Right Justify
         DataField       =   "EMPGROSS"
         DataSource      =   "Data5"
         Height          =   375
         Left            =   -73800
         MaxLength       =   10
         TabIndex        =   27
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtBasic 
         DataField       =   "EMPBASIC"
         DataSource      =   "Data5"
         Height          =   375
         Left            =   -73800
         TabIndex        =   26
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtHRA 
         DataField       =   "EMPHRA"
         DataSource      =   "Data5"
         Height          =   375
         Left            =   -73800
         TabIndex        =   25
         Top             =   1740
         Width           =   1335
      End
      Begin VB.TextBox txtCALL 
         DataField       =   "EMPCALL"
         DataSource      =   "Data5"
         Height          =   375
         Left            =   -73800
         TabIndex        =   24
         Top             =   2310
         Width           =   1335
      End
      Begin VB.TextBox txtUALL 
         DataField       =   "EMPUALL"
         DataSource      =   "Data5"
         Height          =   375
         Left            =   -73800
         TabIndex        =   23
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtDOB 
         DataField       =   "EMPDOB"
         DataSource      =   "Data2"
         Height          =   315
         Left            =   1320
         TabIndex        =   16
         Top             =   2160
         Width           =   1335
      End
      Begin VB.ComboBox cmbSex 
         DataField       =   "EMPSEX"
         DataSource      =   "Data2"
         Height          =   315
         ItemData        =   "frmEmpRecAdd.frx":008C
         Left            =   4560
         List            =   "frmEmpRecAdd.frx":0096
         TabIndex        =   15
         Text            =   "Male"
         Top             =   1080
         Width           =   975
      End
      Begin VB.ComboBox cmbComp 
         DataField       =   "EMPCOMP"
         DataSource      =   "Data4"
         Height          =   315
         Left            =   4440
         TabIndex        =   13
         Top             =   3120
         Width           =   1935
      End
      Begin VB.ComboBox cmbDesg 
         DataField       =   "EMPDESG"
         DataSource      =   "Data3"
         Height          =   315
         ItemData        =   "frmEmpRecAdd.frx":00A8
         Left            =   4440
         List            =   "frmEmpRecAdd.frx":00AA
         TabIndex        =   11
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox txtDOJ 
         DataField       =   "EMPDOJ"
         DataSource      =   "Data2"
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtBLG 
         DataField       =   "EMPBLG"
         DataSource      =   "Data2"
         Height          =   315
         Left            =   4440
         TabIndex        =   8
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtFathName 
         DataField       =   "EMPFATH"
         DataSource      =   "Data2"
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   1635
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         DataField       =   "EMPNAME"
         DataSource      =   "Data2"
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtCode 
         DataField       =   "EMPCODE"
         DataSource      =   "Data2"
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label29 
         Caption         =   "Earned Leaved"
         Height          =   375
         Left            =   -74640
         TabIndex        =   62
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label28 
         Caption         =   "Casual Leaves"
         Height          =   375
         Left            =   -74640
         TabIndex        =   60
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label27 
         Caption         =   "P.F.Code No."
         Height          =   255
         Left            =   -74640
         TabIndex        =   58
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label26 
         Caption         =   "E.S.I. Code No."
         Height          =   255
         Left            =   -74640
         TabIndex        =   56
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label25 
         Caption         =   "Nom Rel."
         Height          =   285
         Left            =   -70440
         TabIndex        =   54
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Nominee Name"
         Height          =   285
         Left            =   -74640
         TabIndex        =   52
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Permanent Tel./PP."
         Height          =   285
         Left            =   -74640
         TabIndex        =   50
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Permanent Address"
         Height          =   285
         Left            =   -74760
         TabIndex        =   47
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Present Tel."
         Height          =   285
         Left            =   -74520
         TabIndex        =   45
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Present Address"
         Height          =   285
         Left            =   -74640
         TabIndex        =   42
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Spl.All."
         Height          =   375
         Left            =   -72360
         TabIndex        =   36
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "L.T.A."
         Height          =   375
         Left            =   -72360
         TabIndex        =   35
         Top             =   2310
         Width           =   975
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Medical"
         Height          =   375
         Left            =   -72360
         TabIndex        =   34
         Top             =   1740
         Width           =   975
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "D.A."
         Height          =   375
         Left            =   -72360
         TabIndex        =   33
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Gross"
         Height          =   375
         Left            =   -74880
         TabIndex        =   32
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Basic"
         Height          =   375
         Left            =   -74880
         TabIndex        =   31
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "H.R.A"
         Height          =   375
         Left            =   -74880
         TabIndex        =   30
         Top             =   1740
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Conv.All."
         Height          =   375
         Left            =   -74880
         TabIndex        =   29
         Top             =   2310
         Width           =   975
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Uniform All."
         Height          =   375
         Left            =   -74880
         TabIndex        =   28
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Code"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Name"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Fath Name"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1605
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Date of Joining"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   2610
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Department"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Date of Birth"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2115
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Sex"
         Height          =   315
         Left            =   3960
         TabIndex        =   14
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Company "
         Height          =   315
         Left            =   3360
         TabIndex        =   12
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Designation"
         Height          =   315
         Left            =   3120
         TabIndex        =   10
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Blood Group"
         Height          =   315
         Left            =   3120
         TabIndex        =   7
         Top             =   2160
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmAddMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Select Case strVal
    Case "M"
    Case "A"
'        Data2.UpdateRecord
    End Select
End Sub

Private Sub Command3_Click()
    frmEmpRec.Show
    frmEmpRec.SetFocus
    Unload frmAddMod
End Sub

Private Sub Data2_Reposition()
    If strVal = "M" Then
    Dim strSql As String
    txtCode.Text = Data2.Recordset.Fields(0)
    Data5.Refresh
    strSql = "Select * from mastsal " _
    & "where mastsal.empcode = '" & txtCode.Text & "'"
    Data5.RecordSource = strSql
    Data5.Refresh
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strSql As String
    Data1.Refresh
    Data2.Refresh
    Data3.Refresh
    Data4.Refresh
    Data5.Refresh
    
    Data1.Recordset.MoveLast
    Data1.Recordset.MoveFirst
    For i = 1 To Data1.Recordset.RecordCount
        Combo1.AddItem (Data1.Recordset.Fields(1))
        Data1.Recordset.MoveNext
    Next
    
    Data3.Refresh
    Data3.Recordset.MoveLast
    Data3.Recordset.MoveFirst
    For i = 1 To Data3.Recordset.RecordCount
        cmbDesg.AddItem (Data3.Recordset.Fields(1))
        Data3.Recordset.MoveNext
    Next
    
    Data4.Refresh
    Data4.Recordset.MoveLast
    Data4.Recordset.MoveFirst
    For i = 1 To Data4.Recordset.RecordCount
        cmbComp.AddItem (Data4.Recordset.Fields(1))
        Data4.Recordset.MoveNext
    Next
    
    txtCode.Enabled = False
    
    Select Case strVal
    Case "M"
        txtCode.Text = Data2.Recordset.Fields(0)
        strSql = "Select * from mastsal " _
        & "where mastsal.empcode = '" & txtCode.Text & "'"
        Data5.RecordSource = strSql
        Data5.Refresh
    Case "A"
        Data2.Recordset.AddNew
        Data5.RecordSource = "MastSal"
        Data5.Refresh
        Data5.Recordset.AddNew
    End Select
End Sub
