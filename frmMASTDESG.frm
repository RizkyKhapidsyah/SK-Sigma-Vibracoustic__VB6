VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMASTDESG 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Designation"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   LinkTopic       =   "Designation "
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmMASTDESG.frx":0000
      Height          =   4335
      Left            =   2760
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "EmpDesg"
         Caption         =   "EmpDesg"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "DesgDes"
         Caption         =   "DesgDes"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
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
      RecordSource    =   "Select EmpDesg, DesgDes from MastDesg"
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   570
      TabIndex        =   4
      Top             =   720
      Width           =   1395
   End
End
Attribute VB_Name = "frmMASTDESG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
Dim iMax As Long
If Text1.Text <> "" Then
   With Adodc1.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         iMax = 0
         Do While Not .EOF
            If iMax < .Fields("EmpDesg") Then
               iMax = .Fields("EmpDesg")
            End If
            .MoveNext
         Loop
      End If
      iMax = iMax + 1
      .AddNew
      !empdesg = iMax
      !desgdes = Trim(Text1)
      .Update
      Text1 = ""
   End With
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

