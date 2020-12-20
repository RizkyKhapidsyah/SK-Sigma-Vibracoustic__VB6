VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4140
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7245
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Index           =   1
      Interval        =   3500
      Left            =   3120
      Top             =   1800
   End
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7080
      Begin VB.Image imgLogo 
         Height          =   1785
         Left            =   120
         Picture         =   "frmSplash.frx":0000
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   2895
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright : Harbir Waan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   3
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Caption         =   "Company : Sigma Vibracoustic (I) Pvt Ltd."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   2
         Top             =   3360
         Width           =   3015
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version 4.0"
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
         Left            =   5580
         TabIndex        =   4
         Top             =   2700
         Width           =   1275
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Platform : Windows 2000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3105
         TabIndex        =   5
         Top             =   2340
         Width           =   3750
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Sigma Vibracoustic"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   600
         TabIndex        =   7
         Top             =   1140
         Width           =   5985
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "LicenseTo Sigma Vibracoustic"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Personnel Management System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1560
         TabIndex        =   6
         Top             =   705
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.Title
End Sub

Private Sub Frame1_Click()
  
    Unload Me
End Sub

Private Sub timer1_Timer(Index As Integer)
If Timer1.UBound Then
Unload Me
End If
End Sub

