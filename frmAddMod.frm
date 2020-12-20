VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddMod 
   Caption         =   "Add Employee Record"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7170
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4800
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.CommandButton Command3 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2400
      TabIndex        =   34
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   33
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Height          =   375
      Left            =   240
      TabIndex        =   32
      Top             =   4800
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   7
      Tab             =   1
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "Emp.Det."
      TabPicture(0)   =   "frmAddMod.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtCode"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtName"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtFathName"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtBLG"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmbDesg"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmbComp"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Cmbdept"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "chkSex"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtDOB"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtDOJ"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Address"
      TabPicture(1)   =   "frmAddMod.frx":001C
      Tab(1).ControlEnabled=   -1  'True
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
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtCA2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtCATel"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtPA1"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtPA2"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtPATel"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtNom"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtNomRel"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Salary"
      TabPicture(2)   =   "frmAddMod.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label15"
      Tab(2).Control(1)=   "Label14"
      Tab(2).Control(2)=   "Label13"
      Tab(2).Control(3)=   "Label12"
      Tab(2).Control(4)=   "Label11"
      Tab(2).Control(5)=   "Label16"
      Tab(2).Control(6)=   "Label17"
      Tab(2).Control(7)=   "Label18"
      Tab(2).Control(8)=   "Label19"
      Tab(2).Control(9)=   "txtUALL"
      Tab(2).Control(10)=   "txtCALL"
      Tab(2).Control(11)=   "txtHRA"
      Tab(2).Control(12)=   "txtBasic"
      Tab(2).Control(13)=   "txtGross"
      Tab(2).Control(14)=   "txtDA"
      Tab(2).Control(15)=   "txtMedical"
      Tab(2).Control(16)=   "txtLTA"
      Tab(2).Control(17)=   "txtSPLALL"
      Tab(2).ControlCount=   18
      TabCaption(3)   =   "ESI && PF"
      TabPicture(3)   =   "frmAddMod.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label26"
      Tab(3).Control(1)=   "Label27"
      Tab(3).Control(2)=   "txtESICode"
      Tab(3).Control(3)=   "txtPFCode"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Leaves"
      TabPicture(4)   =   "frmAddMod.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label28"
      Tab(4).Control(1)=   "Label29"
      Tab(4).Control(2)=   "txtCL"
      Tab(4).Control(3)=   "txtEL"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Resign"
      TabPicture(5)   =   "frmAddMod.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label9"
      Tab(5).Control(1)=   "Label30"
      Tab(5).Control(2)=   "txtDOR"
      Tab(5).Control(3)=   "chkResign"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "Quali / Exp "
      TabPicture(6)   =   "frmAddMod.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame1"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Frame2"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).ControlCount=   2
      Begin VB.Frame Frame2 
         Caption         =   "Experiance"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   68
         Top             =   2400
         Width           =   6855
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   1200
            TabIndex        =   78
            Top             =   1440
            Width           =   5535
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   1200
            TabIndex        =   77
            Top             =   1080
            Width           =   5535
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   1200
            TabIndex        =   76
            Top             =   720
            Width           =   5535
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   1200
            TabIndex        =   75
            Top             =   360
            Width           =   5535
         End
         Begin VB.Label Label37 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "4."
            Height          =   255
            Left            =   600
            TabIndex        =   82
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label Label36 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "3."
            Height          =   255
            Left            =   600
            TabIndex        =   81
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label35 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2."
            Height          =   255
            Left            =   600
            TabIndex        =   80
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label34 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1."
            Height          =   255
            Left            =   600
            TabIndex        =   79
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Qualification"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   67
         Top             =   360
         Width           =   6855
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1200
            TabIndex        =   74
            Top             =   1320
            Width           =   5535
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1200
            TabIndex        =   73
            Top             =   840
            Width           =   5535
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1200
            TabIndex        =   72
            Top             =   360
            Width           =   5535
         End
         Begin VB.Label Label33 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Technical"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label32 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Professional"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label31 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Academic"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.CheckBox chkResign 
         Alignment       =   1  'Right Justify
         Caption         =   "Resigned"
         Height          =   495
         Left            =   -74520
         TabIndex        =   65
         Top             =   720
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker txtDOR 
         Height          =   375
         Left            =   -73080
         TabIndex        =   64
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   24510465
         CurrentDate     =   36370
      End
      Begin MSComCtl2.DTPicker txtDOJ 
         Height          =   315
         Left            =   -73680
         TabIndex        =   7
         Tag             =   "7"
         Top             =   2640
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   36370
      End
      Begin MSComCtl2.DTPicker txtDOB 
         Height          =   315
         Left            =   -73680
         TabIndex        =   5
         Tag             =   "5"
         Top             =   2160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   36370
      End
      Begin VB.CheckBox chkSex 
         Caption         =   "Male"
         Height          =   375
         Left            =   -70560
         TabIndex        =   3
         Tag             =   "3"
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox txtEL 
         Height          =   375
         Left            =   -73200
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtCL 
         Height          =   375
         Left            =   -73200
         TabIndex        =   11
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtPFCode 
         Height          =   285
         Left            =   -72960
         TabIndex        =   14
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtESICode 
         Height          =   285
         Left            =   -72960
         TabIndex        =   13
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtNomRel 
         DataField       =   "EMPNOMREL"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   2040
         TabIndex        =   22
         Top             =   3480
         Width           =   2415
      End
      Begin VB.TextBox txtNom 
         DataField       =   "EMPNOM"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   2040
         TabIndex        =   21
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox txtPATel 
         DataField       =   "EMPPATEL"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   2040
         TabIndex        =   20
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox txtPA2 
         DataField       =   "EMPPA2"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   2040
         TabIndex        =   19
         Top             =   2160
         Width           =   3615
      End
      Begin VB.TextBox txtPA1 
         DataField       =   "EMPPA1"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   2040
         TabIndex        =   18
         Top             =   1800
         Width           =   3615
      End
      Begin VB.TextBox txtCATel 
         DataField       =   "EMPCATEL"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtCA2 
         DataField       =   "EMPCA2"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   2040
         TabIndex        =   16
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtCA1 
         DataField       =   "EMPCA1"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   2040
         TabIndex        =   15
         Top             =   600
         Width           =   3495
      End
      Begin VB.ComboBox Cmbdept 
         Height          =   315
         Left            =   -73680
         TabIndex        =   9
         Tag             =   "9"
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox txtSPLALL 
         Height          =   375
         Left            =   -71160
         MaxLength       =   8
         TabIndex        =   31
         Text            =   "0"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtLTA 
         Height          =   375
         Left            =   -71160
         MaxLength       =   8
         TabIndex        =   29
         Text            =   "0"
         Top             =   2310
         Width           =   1335
      End
      Begin VB.TextBox txtMedical 
         Height          =   375
         Left            =   -71160
         MaxLength       =   12
         TabIndex        =   27
         Text            =   "0"
         Top             =   1740
         Width           =   1335
      End
      Begin VB.TextBox txtDA 
         Height          =   375
         Left            =   -71160
         MaxLength       =   12
         TabIndex        =   25
         Text            =   "0"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtGross 
         Height          =   375
         Left            =   -73800
         MaxLength       =   12
         TabIndex        =   23
         Text            =   "0"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtBasic 
         Height          =   375
         Left            =   -73800
         MaxLength       =   12
         TabIndex        =   24
         Text            =   "0"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtHRA 
         Height          =   375
         Left            =   -73800
         MaxLength       =   12
         TabIndex        =   26
         Text            =   "0"
         Top             =   1740
         Width           =   1335
      End
      Begin VB.TextBox txtCALL 
         Height          =   375
         Left            =   -73800
         MaxLength       =   8
         TabIndex        =   28
         Text            =   "0"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtUALL 
         Height          =   375
         Left            =   -73800
         MaxLength       =   8
         TabIndex        =   30
         Text            =   "0"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ComboBox cmbComp 
         Height          =   315
         Left            =   -70560
         TabIndex        =   10
         Tag             =   "10"
         Top             =   3120
         Width           =   1935
      End
      Begin VB.ComboBox cmbDesg 
         Height          =   315
         Left            =   -70560
         TabIndex        =   8
         Tag             =   "8"
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox txtBLG 
         Height          =   315
         Left            =   -70560
         TabIndex        =   6
         Tag             =   "6"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtFathName 
         Height          =   315
         Left            =   -73680
         TabIndex        =   4
         Tag             =   "4"
         Top             =   1635
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   -73680
         TabIndex        =   2
         Tag             =   "2"
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtCode 
         Height          =   315
         Left            =   -73680
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label30 
         Caption         =   "Note :-   Resign Date should be employees last working day in the Organisation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74520
         TabIndex        =   66
         Top             =   2280
         Width           =   4335
      End
      Begin VB.Label Label9 
         Caption         =   "Resignation Date"
         Height          =   375
         Left            =   -74520
         TabIndex        =   63
         Top             =   1440
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
         TabIndex        =   61
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label27 
         Caption         =   "P.F.Code No."
         Height          =   255
         Left            =   -74640
         TabIndex        =   60
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label26 
         Caption         =   "E.S.I. Code No."
         Height          =   255
         Left            =   -74640
         TabIndex        =   59
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label25 
         Caption         =   "Nom Rel."
         Height          =   285
         Left            =   960
         TabIndex        =   58
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Nominee Name"
         Height          =   285
         Left            =   360
         TabIndex        =   57
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Permanent Tel./PP."
         Height          =   285
         Left            =   360
         TabIndex        =   56
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Permanent Address"
         Height          =   285
         Left            =   240
         TabIndex        =   55
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Present Tel."
         Height          =   285
         Left            =   480
         TabIndex        =   54
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Present Address"
         Height          =   285
         Left            =   360
         TabIndex        =   53
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Spl.All."
         Height          =   375
         Left            =   -72360
         TabIndex        =   52
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "L.T.A."
         Height          =   375
         Left            =   -72360
         TabIndex        =   51
         Top             =   2310
         Width           =   975
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Medical"
         Height          =   375
         Left            =   -72360
         TabIndex        =   50
         Top             =   1740
         Width           =   975
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "D.A."
         Height          =   375
         Left            =   -72360
         TabIndex        =   49
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Gross"
         Height          =   375
         Left            =   -74880
         TabIndex        =   48
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Basic"
         Height          =   375
         Left            =   -74880
         TabIndex        =   47
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "H.R.A"
         Height          =   375
         Left            =   -74880
         TabIndex        =   46
         Top             =   1740
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Conv.All."
         Height          =   375
         Left            =   -74880
         TabIndex        =   45
         Top             =   2310
         Width           =   975
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Uniform All."
         Height          =   375
         Left            =   -74880
         TabIndex        =   44
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Code"
         Height          =   375
         Index           =   0
         Left            =   -74880
         TabIndex        =   43
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Name"
         Height          =   375
         Left            =   -74880
         TabIndex        =   42
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Fath Name"
         Height          =   375
         Left            =   -74880
         TabIndex        =   41
         Top             =   1605
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Date of Joining"
         Height          =   375
         Left            =   -74880
         TabIndex        =   40
         Top             =   2610
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Department"
         Height          =   375
         Left            =   -74880
         TabIndex        =   39
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Date of Birth"
         Height          =   375
         Left            =   -74880
         TabIndex        =   38
         Top             =   2115
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Company "
         Height          =   315
         Left            =   -71640
         TabIndex        =   37
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Designation"
         Height          =   315
         Left            =   -71880
         TabIndex        =   36
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Blood Group"
         Height          =   315
         Left            =   -71880
         TabIndex        =   35
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
Dim dbpms As Database
Dim rstPM, rstPS, rstDept, rstDesg, rstComp, rstSer As Recordset

Private Sub chkResign_Click()
If chkResign = 1 Then
   txtDOR.Enabled = True
   txtDOR.MinDate = CDate(Str(Month(Now - 1)) & "/01/" & Str(Year(Now - 1)))
   txtDOR.MaxDate = CDate(Now)
Else
   txtDOR.Enabled = False
End If
End Sub

Private Sub Command1_Click()
Dim intSer As Integer
Dim strNewCode, strSql As String
Dim intLen As Integer
cmbupdate
If strVal = "A" Then
    With rstSer
        .MoveFirst
        .Edit
        intSer = !empser
        intSer = intSer + 1
        !empser = intSer
        intLen = Len(Trim(CStr(intSer)))
        Select Case intLen
        Case 1:
            strNewCode = "A0000" & Trim(CStr(intSer))
        Case 2:
            strNewCode = "A000" & Trim(CStr(intSer))
        Case 3:
            strNewCode = "A00" & Trim(CStr(intSer))
        Case 4:
            strNewCode = "A0" & Trim(CStr(intSer))
        Case 5:
            strNewCode = Trim(CStr(intSer))
        End Select
        txtCode = strNewCode
        .Update
    End With
End If

With rstPM

    If strVal = "A" Then
        .AddNew
        !EMPCODE = txtCode
        !EMPNAME = txtName
        !empsex = chkSex.Value
        !empdob = CDate(txtDOB)
        !EMPDOJ = CDate(txtDOJ)
        !empblg = txtBLG
        !empfath = txtFathName
        !empdept = Cmbdept.Text
        !empdesg = cmbDesg
        !empcomp = cmbComp
        !emppa1 = txtPA1
        !emppa2 = txtPA2
        !emppatel = txtPATel
        !empca1 = txtCA1
        !empca2 = txtCA2
        !empcatel = txtCATel
        !empnom = txtNom
        !empnomrel = txtNomRel
        !emppf = txtPFCode
        !empesi = txtESICode
        !empcltot = Val(txtCL)
        !empeltot = Val(txtEL)
        If chkResign = 1 Then
            !EMPDOR = txtDOR
         Else
            !EMPDOR = "01/01/1900"
         End If
        .Update
    Else
        .FindFirst "Empcode = '" & txtCode & "'"
        If Not .EOF Then
        .Edit
        !EMPCODE = txtCode
        !EMPNAME = txtName
        !empsex = chkSex.Value
        !empdob = txtDOB
        !EMPDOJ = txtDOJ
        !empblg = txtBLG
        !empfath = txtFathName
        !empdept = Cmbdept.Text
        !empdesg = cmbDesg
        !empcomp = cmbComp
        !emppa1 = txtPA1
        !emppa2 = txtPA2
        !emppatel = txtPATel
        !empca1 = txtCA1
        !empca2 = txtCA2
        !empcatel = txtCATel
        !empnom = txtNom
        !empnomrel = txtNomRel
        !emppf = txtPFCode
        !empesi = txtESICode
        !empcltot = txtCL
        !empeltot = txtEL
         If chkResign = 1 Then
            !EMPDOR = txtDOR
         Else
            !EMPDOR = "01/01/1900"
         End If
        .Update
        End If
    End If
End With
    
With rstPS
    If strVal = "A" Then
        .AddNew
        !EMPCODE = txtCode
        !empbasic = txtBasic
        !empda = txtDA
        !emphra = txtHRA
        !empmed = txtMedical
        !empcall = txtCALL
        !empltall = txtLTA
        !empuall = txtUALL
        !empsplall = txtSPLALL
        .Update
    Else
        .FindFirst "Empcode = '" & txtCode & "'"
        If Not .EOF Then
        .Edit
        !EMPCODE = txtCode
        !empbasic = txtBasic
        !empda = txtDA
        !emphra = txtHRA
        !empmed = txtMedical
        !empcall = txtCALL
        !empltall = txtLTA
        !empuall = txtUALL
        !empsplall = txtSPLALL
        .Update
        End If
    End If
End With
Command3_Click
End Sub

Private Sub Command3_Click()
    Unload frmAddMod
    frmMain.Show
End Sub



Private Sub Form_Load()

On Error GoTo ErrorHandler
Dim i As Integer
Dim strSql As String
    
Dim mgross As Single
Dim mbasic As Single
Dim mda As Single
Dim mhra As Single
Dim mmed As Single
Dim mca As Single
Dim mlta As Single
Dim mua As Single
Dim mspl As Single
    
txtGross.Text = 0
txtBasic.Text = 0
txtDA.Text = 0
txtHRA.Text = 0
txtMedical.Text = 0
txtCALL.Text = 0
txtLTA.Text = 0
txtUALL.Text = 0
txtSPLALL.Text = 0
  
Set dbpms = OpenDatabase(strPath & "dbpms.mdb")
Set rstPM = dbpms.OpenRecordset("mastpay", dbOpenDynaset)
Set rstPS = dbpms.OpenRecordset("mastsal", dbOpenDynaset)
Set rstDept = dbpms.OpenRecordset("mastdept", dbOpenDynaset)
Set rstDesg = dbpms.OpenRecordset("mastdesg", dbOpenDynaset)
Set rstComp = dbpms.OpenRecordset("mastcomp", dbOpenDynaset)
Set rstSer = dbpms.OpenRecordset("empcodes", dbOpenDynaset)

SSTab1.Tab = 0
    
With rstDept
.MoveLast
.MoveFirst
For i = 1 To .RecordCount
    Cmbdept.AddItem (!deptdes)
    .MoveNext
Next
End With

With rstDesg
.MoveLast
.MoveFirst
For i = 1 To .RecordCount
    cmbDesg.AddItem (!desgdes)
    .MoveNext
Next
End With

With rstComp
.MoveLast
.MoveFirst
For i = 1 To .RecordCount
    cmbComp.AddItem (!compdes)
    .MoveNext
Next
End With
  
txtCode.Enabled = False
  
Select Case strVal
Case "M"

    strSql = "EmpCode = '" & strVal1 & "'"
    With rstPM
        .FindFirst strSql
        If Not .EOF Then
            txtCode = !EMPCODE
            txtName = !EMPNAME
            txtFathName = !empfath
            chkSex.Value = !empsex
            txtDOB = !empdob
            txtBLG = !empblg
            txtDOJ = !EMPDOJ
            Cmbdept = !empdept
            cmbDesg = !empdesg
            cmbComp = !empcomp
            txtPA1 = !emppa1
            txtPA2 = !emppa2
            txtPATel = !emppatel
            txtCA1 = !empca1
            txtCA2 = !empca2
            txtCATel = !empcatel
            txtNom = !empnom
            txtNomRel = !empnomrel
            txtPFCode = !emppf
            txtESICode = !empesi
            txtCL = !empcltot
            txtEL = !empeltot
            If IsNull(!EMPDOR) = True Or !EMPDOR = CDate("01/01/1900") Then
               txtDOR = Now
               chkResign = 0
               txtDOR.Enabled = False
            Else
               txtDOR.Enabled = True
               chkResign = 1
               txtDOR.MinDate = !EMPDOR
               txtDOR.MaxDate = !EMPDOR
               txtDOR = !EMPDOR
            End If
        End If
    End With

    strSql = "empcode = '" & strVal1 & "'"
    rstPS.FindFirst strSql
    With rstPS
        If Not .EOF Then
            txtBasic = !empbasic
            txtDA = !empda
            txtHRA = !emphra
            txtMedical = !empmed
            txtCALL = !empcall
            txtLTA = !empltall
            txtUALL = !empuall
            txtSPLALL = !empsplall
        Else
            MsgBox "Salary Details Not Found .. Error "
            Close
        End If
    End With
    
    With rstDept
        .FindFirst "Empdept = '" & Cmbdept.Text & "'"
        If Not .EOF Then
            Cmbdept.Text = !deptdes
        Else
            Cmbdept = " "
        End If
    End With
    With rstDesg
        .FindFirst "Empdesg = '" & cmbDesg.Text & "'"
        If Not .EOF Then
            cmbDesg.Text = !desgdes
        Else
            cmbDesg = " "
        End If
    End With
    With rstComp
        .FindFirst "Empcomp = '" & cmbComp.Text & "'"
        If Not .EOF Then
            cmbComp.Text = !compdes
        Else
            cmbComp = " "
        End If
    End With
    
Case "A"
    
End Select
    
mgross = Val(txtGross)
mbasic = Val(txtBasic)
mda = Val(txtDA)
mhra = Val(txtHRA)
mmed = Val(txtMedical)
mca = Val(txtCALL)
mlta = Val(txtLTA)
mua = Val(txtUALL)
mspl = Val(txtSPLALL)

mgross = mbasic + mda + mhra + mmed + mca + mlta + mua + mspl
    
txtGross = Format$(mgross, "#.00")
Exit Sub
ErrorHandler:
MsgBox Err.Number & " " & Err.Description
Unload Me
End Sub

Private Sub cmbupdate()
Dim strSql As String
If Cmbdept.Text <> "" Then
    strSql = "deptdes = '" & Cmbdept.Text & "'"
    With rstDept
        .FindFirst strSql
        If Not .EOF Then
            Cmbdept.Text = !empdept
        Else
            Cmbdept.Text = " "
        End If
    End With
End If
            
If cmbDesg.Text <> " " Then
    strSql = "desgdes = '" & cmbDesg.Text & "'"
    With rstDesg
        .FindFirst strSql
        If Not .EOF Then
            cmbDesg.Text = !empdesg
        Else
            cmbDesg.Text = " "
        End If
    End With
End If

If cmbComp.Text <> " " Then
    strSql = "compdes = '" & cmbComp.Text & "'"
    With rstComp
        .FindFirst strSql
        If Not .EOF Then
            cmbComp.Text = !empcomp
        Else
            cmbComp.Text = " "
        End If
    End With
End If
End Sub

Private Sub totgross(mascii As Integer)
If mascii = 13 Then
Dim mgross As Single
Dim mbasic As Single
Dim mda As Single
Dim mhra As Single
Dim mmed As Single
Dim mca As Single
Dim mlta As Single
Dim mua As Single
Dim mspl As Single

mgross = Val(txtGross)
mbasic = Val(txtBasic)
mda = Val(txtDA)
mhra = Val(txtHRA)
mmed = Val(txtMedical)
mca = Val(txtCALL)
mlta = Val(txtLTA)
mua = Val(txtUALL)
mspl = Val(txtSPLALL)

mgross = mbasic + mda + mhra + mmed + mca + mlta + mua + mspl

txtGross = Format$(mgross, "#.00")
End If
End Sub


Private Sub txtBasic_KeyPress(KeyAscii As Integer)
totgross (KeyAscii)
End Sub
Private Sub txtda_KeyPress(KeyAscii As Integer)
totgross (KeyAscii)
End Sub



Private Sub txthra_KeyPress(KeyAscii As Integer)
totgross (KeyAscii)
End Sub
Private Sub txtMedical_KeyPress(KeyAscii As Integer)
totgross (KeyAscii)
End Sub
Private Sub txtCAll_KeyPress(KeyAscii As Integer)
totgross (KeyAscii)
End Sub
Private Sub txtlta_KeyPress(KeyAscii As Integer)
totgross (KeyAscii)
End Sub
Private Sub txtuall_KeyPress(KeyAscii As Integer)
totgross (KeyAscii)
End Sub
Private Sub txtsplall_KeyPress(KeyAscii As Integer)
totgross (KeyAscii)
End Sub

Private Sub txtGross_KeyPress(KeyAscii As Integer)
If InStr(1, txtGross, "0123456789") Then

End If
If KeyAscii = 13 Then
Dim mgross As Single
Dim mbasic As Single
Dim mda As Single
Dim mhra As Single
Dim mmed As Single
Dim mca As Single
Dim mlta As Single
Dim mua As Single
Dim mspl As Single

mgross = Val(txtGross)
mbasic = Val(txtBasic)
mda = Val(txtDA)
mhra = Val(txtHRA)
mmed = Val(txtMedical)
mca = Val(txtCALL)
mlta = Val(txtLTA)
mua = Val(txtUALL)
mspl = Val(txtSPLALL)

If mgross > 0 And mgross < 9999999 Then
   If mgross <= 2500 Then
      mbasic = mgross * 0.55
      mhra = mbasic * 0.4
      mmed = 0
      mca = Round(mbasic * 0.3, 2)
      mlta = 0
      mua = mgross - mbasic - mhra - mca
      mda = 0
      mspl = 0
   Else
    mbasic = mgross * 0.5
    mhra = mbasic * 0.4
    mmed = mbasic * 0.1
    mca = mbasic * 0.3
    mlta = mbasic * 0.1
    mua = mbasic * 0.1
    mda = 0
    mspl = 0
    End If
    txtBasic = Format$(mbasic, "#.00")
    txtDA = Format$(mda, "#.00")
    txtHRA = Format$(mhra, "#.00")
    txtMedical = Format$(mmed, "#.00")
    txtCALL = Format$(mca, "#.00")
    txtLTA = Format$(mlta, "#.00")
    txtUALL = Format$(mua, "#.00")
    txtSPLALL = Format$(mspl, "#.00")
End If
End If
End Sub

Private Sub Check_KeyPress(KeyAscii As Integer)
MsgBox KeyAscii
End Sub

