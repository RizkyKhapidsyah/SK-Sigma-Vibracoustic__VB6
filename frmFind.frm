VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Search"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Search on"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Find"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   780
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload frmFind
frmMain.Show
End Sub
