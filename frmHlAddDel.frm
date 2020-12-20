VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmHlAddDel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Holiday Addition/Deletion"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3000
      Top             =   720
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
      RecordSource    =   "TRANLEAV"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3000
      Top             =   240
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
   Begin VB.CommandButton Command3 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24510465
      CurrentDate     =   36491
   End
   Begin VB.Label lb2 
      Caption         =   "Label2"
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date of Holiday"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmHlAddDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Dim xEmpCode As String
Adodc1.Refresh
With Adodc1.Recordset
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            xEmpCode = !EMPCODE
            If Format$(dor, "MM-DD-YYYY") <> Format$("01/01/1900", "MM-DD-YYYY") Then
            Adodc2.CommandType = adCmdText
            Adodc2.RecordSource = "SELECT * FROM TRANLEAV WHERE EMPCODE = '" & Trim(xEmpCode) & "' AND YR = " & Year(DTPicker1) & " AND MN = " & Month(DTPicker1) & " AND DY = " & Day(DTPicker1)
            'DOL = #" & Format$(DTPicker1, "MM-DD-YYYY") & "#"
            Adodc2.Refresh
            With Adodc2.Recordset
                If .RecordCount = 0 Then
                    lb2 = xEmpCode & "   " & .RecordCount
                    .AddNew
                    !YR = Year(DTPicker1)
                    !MN = Month(DTPicker1)
                    !DY = Day(DTPicker1)
                    !EMPCODE = xEmpCode
                    !DOL = DTPicker1.Value
                    !LEAVTYPE = "H"
                    !DOF = Date
                    !LEAVDAY = 1
                    .Update
                End If
            End With
            End If
            .MoveNext
        Loop
    End If
End With
Command3.SetFocus
End Sub

Private Sub cmdDel_Click()
Dim xEmpCode As String
Adodc1.Refresh
With Adodc1.Recordset
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            xEmpCode = !EMPCODE
            Adodc2.CommandType = adCmdText
            Adodc2.RecordSource = "SELECT * FROM TRANLEAV WHERE EMPCODE = '" & Trim(xEmpCode) & "' AND DOL = #" & Format$(DTPicker1, "MM-DD-YYYY") & "#"
            Adodc2.Refresh
            With Adodc2.Recordset
                If .RecordCount > 0 Then
                    If !LEAVTYPE = "H" Then
                        .Delete
                        .Update
                    End If
                Else
                End If
            End With
            .MoveNext
        Loop
    End If
End With
End Sub

Private Sub Command3_Click()
Unload Me
End Sub
