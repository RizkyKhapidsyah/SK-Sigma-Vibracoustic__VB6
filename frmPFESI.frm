VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmPFESI 
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   4560
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "D:\PMS\PMSPrj\dlyleave.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24444929
      CurrentDate     =   36375
   End
   Begin VB.Label Label1 
      Caption         =   "Date"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "frmPFESI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim dtSDate As Date
Dim myr As Integer
Dim mmn As Integer
Dim mdd As Integer

dtSDate = DTPicker1.Value
myr = Year(dtSDate)
mmn = Month(dtSDate)
mdd = Day(dtSDate)
CrystalReport1.SelectionFormula = "{TRANLEAV.DOL} = date(" & Str(myr) & "," & Str(mmn) & "," & Str(mdd) & ")"
CrystalReport1.Action = 1

End Sub

