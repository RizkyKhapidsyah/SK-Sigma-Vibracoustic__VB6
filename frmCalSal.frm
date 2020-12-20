VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCalSal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Salary Calculation"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "MANAGER'S"
      Height          =   375
      Left            =   1440
      TabIndex        =   16
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdresign 
      Caption         =   "Resignation"
      Height          =   375
      Left            =   2760
      TabIndex        =   15
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Cmdcash 
      Caption         =   "Cash"
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   4200
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Tag             =   "5"
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Tag             =   "4"
      Top             =   960
      Width           =   975
   End
   Begin VB.ListBox SelList 
      Height          =   645
      Left            =   3360
      TabIndex        =   1
      Tag             =   "6"
      Top             =   960
      Width           =   2175
   End
   Begin VB.ListBox TotList 
      Height          =   645
      Left            =   240
      TabIndex        =   0
      Tag             =   "3"
      Top             =   960
      Width           =   2175
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5280
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "D:\PMS\PMSPrj\salnew.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   7
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Tag             =   "b8"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Bank"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Tag             =   "7"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox txtMonth 
      Height          =   375
      Left            =   960
      MaxLength       =   2
      TabIndex        =   7
      Tag             =   "1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtYear 
      Height          =   375
      Left            =   3360
      MaxLength       =   4
      TabIndex        =   8
      Tag             =   "2"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Status"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label Label5 
      Height          =   375
      Left            =   840
      TabIndex        =   12
      Top             =   4800
      Width           =   3495
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Caption         =   "Order by"
      Height          =   255
      Left            =   4080
      TabIndex        =   11
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Total Items "
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Month"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Year"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmCalSal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbpms As Database
Dim rstPM As Recordset
Dim rstPS As Recordset
Dim rstTL As Recordset
Dim rstTS As Recordset
Dim rstOTS As Recordset
Dim rstMNBL As Recordset
Dim xOYR, xOMN As Integer
Dim mLtESI As Single
Dim xxctr, xpf As Single
Dim mcode, mdesg, mdept, mcomp As String
Dim mbasic, mda, mmed, mhra, mca, mlta, mua, mspl, mgross As Single
Dim mabasic, mada, mamed, mahra, maca, malta, maua, maspl, magross As Single
Dim mx, mx1 As Date
Dim myr, mmn As Integer
Dim mtotdy, mtotwk, mwkdy, mhldy, mcldy, meldy, mabdy, mPrdy As Single
Dim mxdoj, mxdor As Date
Dim mjoin, mresign, xDoIt As Boolean
Dim a As Integer
Dim mesiamt, mpfamt, mesi, mpf, mepf
Dim mpfno As String

Private Sub cmdAdd_Click()
    Dim msel, ListCtr, i As Integer
    Dim boolFlag As Boolean
    boolFlag = True
    msel = TotList.ListIndex
    ListCtr = SelList.ListCount
    For i = 0 To ListCtr
        If SelList.List(i) = TotList.List(msel) Then
            boolFlag = False
            Exit For
        End If
    Next
    If boolFlag Then
        SelList.AddItem (TotList.List(msel))
    End If
End Sub

Private Sub Cmdcash_Click()
Dim intCtrTest As Integer
Dim mxx As Date
If txtMonth.Text = "" Then
    txtMonth.SetFocus
    Exit Sub
End If
If txtYear.Text = "" Then
    txtYear.SetFocus
    Exit Sub
End If
If strVal = "1" Then
    xOMN = Val(txtMonth) - 1
    If xOMN = 0 Then
        xOMN = 12
        xOYR = Val(txtYear) - 1
    Else
        xOYR = Val(txtYear)
    End If
    Set dbpms = OpenDatabase(strPath & "dbpms.mdb")
    Set rstPM = dbpms.OpenRecordset("select * from mastpay", dbOpenDynaset)
    Set rstPS = dbpms.OpenRecordset("select * from mastsal order by empcode", dbOpenDynaset)
    Set rstTS = dbpms.OpenRecordset("select * from TRANSAL where yr = " & Val(txtYear.Text) & " and mn = " & Val(txtMonth))
    Set rstTL = dbpms.OpenRecordset("select * from TRANLEAV", dbOpenDynaset)
    Set rstOTS = dbpms.OpenRecordset("select * from TRANSAL where yr = " & xOYR & " and mn = " & xOMN)
    Set rstMNBL = dbpms.OpenRecordset("select * from mnblock", dbOpenDynaset)
    myr = txtYear
    mmn = txtMonth
    
    With rstMNBL
      .MoveFirst
      Do While Not .EOF
         If !yrblock = myr And !mnblock = mmn Then
            MsgBox "Unable to Calculate. Month Blocked "
            Exit Sub
         End If
         .MoveNext
      Loop
    End With
    mx = txtMonth & "/01/" & txtYear
    mx1 = DateAdd("m", 1, mx)
    mx1 = DateAdd("d", -1, mx1)
    mxx = mx
    mtotdy = Day(mx1)
    If myr >= 1999 And (mmn >= 1 And mmn <= 12) Then
        With rstTS
            .FindFirst "yr = " & myr & " and mn = " & mmn
            If Not .EOF Then
                If vbYes = MsgBox("Already processed. Process again  ", vbYesNo) Then
                    a = 1
                Else
                    a = 0
                End If
            Else
                a = 1
            End If
        End With
        If a = 1 Then
            With rstTS
                
                If .RecordCount > 0 Then
                While Not .EOF
                    .Delete
                    .MoveNext
                Wend
                End If
            End With
            
            With rstPS
                If .RecordCount > 0 Then
                .MoveFirst
                While Not .EOF
                    xDoIt = True
                    mjoin = False
                    mresign = False
                    mpfno = ""
                    mtotwk = 0
                    mwkdy = 0
                    mhldy = 0
                    mcldy = 0
                    meldy = 0
                    mabdy = 0
                    mtotwk = mtotdy
                    mcode = !EMPCODE
                    mbasic = !empbasic
                    mda = !empda
                    mhra = !emphra
                    mmed = !empmed
                    mca = !empcall
                    mlta = !empltall
                    mua = !empuall
                    mspl = !empsplall
                    
                    mabasic = mbasic
                    mada = mda
                    mahra = mhra
                    mamed = mmed
                    maca = mca
                    malta = mlta
                    maua = mua
                    maspl = mspl
                    Label5.Caption = "Processing Code " & mcode
                    With rstPM
                        .FindFirst "empcode = '" & mcode & "'"
                        If Not .EOF Then
                            mdesg = !empdesg
                            mdept = !empdept
                            mcomp = !empcomp
                            mxdoj = !EMPDOJ
                            mpfno = !emppf
                            If IsNull(!EMPDOR) Then
                            mxdor = CDate("01/01/1900")
                            Else
                            mxdor = !EMPDOR
                            End If
                            If Str(Year(mxdoj)) = Str(myr) And Str(Month(mxdoj)) = Str(mmn) Then
                                mjoin = True
                            Else
                                mjoin = False
                            End If
                            If Str(Year(mxdor)) = Str(myr) And Str(Month(mxdor)) = Str(mmn) Then
                                mresign = True
                                xDoIt = True
                            Else
                                If mxdor = CDate("01/01/1900") Then
                                    mresign = False
                                    xDoIt = True
                                Else
                                    xDoIt = False
                                End If
                            End If
                        Else
                            mdesg = " "
                            mdept = " "
                            mcomp = " "
                            mjoin = False
                            mresign = False
                            xDoIt = False
                        End If
                    End With
                    If xDoIt Then
                    mx = mxx
                    While mx <= mx1
                        If mjoin Then
                            If mx < mxdoj Then
                                mabdy = mabdy + 1
                            Else
                                If Weekday(mx) = 1 Then
                                    mhldy = mhldy + 1
                                End If
                            End If
                        Else
                            If mresign Then
                                If mx > mxdor Then
                                    mabdy = mabdy + 1
                                Else
                                    If Weekday(mx) = 1 Then
                                        mhldy = mhldy + 1
                                    End If
                                End If
                            Else
                                If Weekday(mx) = 1 Then
                                    mhldy = mhldy + 1
                                End If
                            End If
                        End If
                        mx = DateAdd("d", 1, mx)
                    Wend
                    With rstTL
                        .FindFirst "empcode = '" & mcode & "' and yr = " & myr & " and mn = " & mmn
                        intCtrTest = 0
                        While Not .NoMatch
                            intCtrTest = intCtrTest + 1
                            Label5.Caption = "Processing Code " & mcode & "  Leave Processing " & intCtrTest
                            If Weekday(!DOL) = 1 Then
                                mhldy = mhldy - 1
                              End If
                            Select Case !LEAVTYPE
                            Case "H"
                                mhldy = mhldy + !LEAVDAY
                            Case "C"
                                mcldy = mcldy + !LEAVDAY
                            Case "E"
                                meldy = meldy + !LEAVDAY
                            Case "A"
                                mabdy = mabdy + !LEAVDAY
                            Case "L"
                                mabdy = mabdy + !LEAVDAY
                            Case "P"
                                mPrdy = mPrdy + !LEAVDAY
                            End Select
                      
                            .FindNext "empcode = '" & mcode & "' and yr = " & myr & " and mn = " & mmn
                        Wend
                    End With
                    mtotwk = mtotwk - mabdy
                    mwkdy = mtotwk - mhldy - mcldy - meldy
                    If mtotdy <> mtotwk Then
                        mbasic = mbasic * (mtotwk / mtotdy)
                        mda = mda * (mtotwk / mtotdy)
                        mhra = mhra * (mtotwk / mtotdy)
                        mmed = mmed * (mtotwk / mtotdy)
                        mlta = mlta * (mtotwk / mtotdy)
                        mca = mca * (mtotwk / mtotdy)
                        mua = mua * (mtotwk / mtotdy)
                        mspl = mspl * (mtotwk / mtotdy)
                    End If
                    mpfamt = 0
                    mesiamt = 0
                    mpf = 0
                    mepf = 0
                    mesi = 0
                    xpf = 0
                    With rstOTS
                        .FindFirst "empcode = '" & Trim(mcode) & "'"
                        If Not .NoMatch Then
                            mLtESI = !empesi
                        Else
                            mLtESI = 0
                        End If
                    End With
                    
                    If mpfno <> "" Or (mabasic + mada) <= 5000 Then
                        mbasic = Round(mbasic, 2)
                        If mbasic - Int(mbasic) >= 0.5 Then
                            mbasic = Int(mbasic) + 1
                        End If
                        mpfamt = mbasic + mda
                        xpf = ((mbasic + mda) * (12 / 100))
                        If xpf - Int(xpf) >= 0.5 Then
                            xpf = Int(xpf) + 1
                        Else
                            xpf = Int(xpf)
                        End If
                        mpf = (xpf * (3.67 / 100))
                        mpf = Int(mpf)
                        mepf = xpf - mpf
                        If (mepf + mpf) > 780 Then
                            mepf = 541 '417'
                            mpf = 239
                        End If
                    End If
                    If (mabasic + mada + mahra + maua + maspl) <= 6500 Then
                        mbasic = Round(mbasic, 2)
                        If mbasic - Int(mbasic) >= 0.5 Then
                            mbasic = Int(mbasic) + 1
                        End If
                        mesiamt = mbasic + mda + mhra + mua + mspl
                        mesi = (mbasic + mda + mhra + mua + mspl) * (1.75 / 100)
                        mesi = Round(mesi, 2)
                        xxctr = Val(Format$(mesi - Int(mesi), "#.00"))
                        If xxctr > 0 Then
                        xxctr = xxctr * 100
                        If (xxctr Mod 5) = 0 Then
                        Else
                        Do While (xxctr Mod 5) <> 0
                            xxctr = xxctr + 1
                        Loop
                        xxctr = (xxctr / 100)
                        mesi = Int(mesi) + xxctr
                        mesi = Round(mesi, 2)
                        End If
                        End If
                        'If Format$(Int(mesi), "#.00") <> Format$(mesi, "#.00") Then
                        '    mesi = mesi + 1
                        '    mesi = Int(mesi)
                        'End If
                    Else
                        If Val(txtMonth) = 10 Or Val(txtMonth) = 4 Then
                            mesi = 0
                        Else
                            If mLtESI > 0 Then
                            mesi = 113.75
                            Else
                            mesi = 0
                            End If
                        End If
                    End If
                    
                    With rstTS
                        .AddNew
                        !EMPCODE = mcode
                        !YR = myr
                        !MN = mmn
                        !empcomp = mcomp
                        !empdept = mdept
                        !empdesg = mdesg
                        !empbasic = mbasic
                        !empda = mda
                        !emphra = mhra
                        !empmed = mmed
                        !empcall = mca
                        !empltall = mlta
                        !empuall = mua
                        !empsplall = mspl
                        !actesiamt = mesiamt
                        !actpfamt = mpfamt
                        !empesi = mesi
                        !emppf = mpf
                        !empepf = mepf
                        !actbasic = mabasic
                        !actda = mada
                        !acthra = mahra
                        !actmed = mamed
                        !actcall = maca
                        !actltall = malta
                        !actuall = maua
                        !actsplall = maspl
                        !empwkdy = Format$(mwkdy, "#.00")
                        !emphldy = Format$(mhldy, "#.00")
                        !empcldy = Format$(mcldy, "#.00")
                        !empeldy = Format$(meldy, "#.00")
                        !empabdy = Format$(mabdy, "#.00")
                        .Update
                    End With
                    End If
                    .MoveNext
                Wend
                End If
            End With
        Else
        
        End If
    Else
        MsgBox "Invalid Month and Year "
        txtMonth.SetFocus
    End If
End If
If strVal = "2" Then
    CrystalReport1.WindowState = crptMaximized
'    CrystalReport1.ReportFileName = strPath & "salnewfor.rpt"
    CrystalReport1.ReportFileName = strPath & "salCash.rpt"
    CrystalReport1.SelectionFormula = "{transal.yr} = " & Val(Trim(txtYear)) & " and {transal.mn} = " & Val(Trim(txtMonth))
    CrystalReport1.Connect = "DSN = PMS"
'    CrystalReport1.Destination = crptToFile
'    CrystalReport1.PrinterStartPage = 1
'    CrystalReport1.PrinterStopPage = 1
    CrystalReport1.Action = 1
    
End If
End Sub

Private Sub cmdOk_Click()
Dim intCtrTest As Integer
Dim mxx As Date
If txtMonth.Text = "" Then
    txtMonth.SetFocus
    Exit Sub
End If
If txtYear.Text = "" Then
    txtYear.SetFocus
    Exit Sub
End If
If strVal = "1" Then
    xOMN = Val(txtMonth) - 1
    If xOMN = 0 Then
        xOMN = 12
        xOYR = Val(txtYear) - 1
    Else
        xOYR = Val(txtYear)
    End If
    Set dbpms = OpenDatabase(strPath & "dbpms.mdb")
    Set rstPM = dbpms.OpenRecordset("select * from mastpay", dbOpenDynaset)
    Set rstPS = dbpms.OpenRecordset("select * from mastsal order by empcode", dbOpenDynaset)
    Set rstTS = dbpms.OpenRecordset("select * from TRANSAL where yr = " & Val(txtYear.Text) & " and mn = " & Val(txtMonth))
    Set rstTL = dbpms.OpenRecordset("select * from TRANLEAV", dbOpenDynaset)
    Set rstOTS = dbpms.OpenRecordset("select * from TRANSAL where yr = " & xOYR & " and mn = " & xOMN)
    Set rstMNBL = dbpms.OpenRecordset("select * from mnblock", dbOpenDynaset)
    myr = txtYear
    mmn = txtMonth
    
    With rstMNBL
      '.MoveFirst
      Do While Not .EOF
         If !yrblock = myr And !mnblock = mmn Then
            MsgBox "Unable to Calculate. Month Blocked "
            Exit Sub
         End If
         .MoveNext
      Loop
    End With
    mx = txtMonth & "/01/" & txtYear
    mx1 = DateAdd("m", 1, mx)
    mx1 = DateAdd("d", -1, mx1)
    mxx = mx
    mtotdy = Day(mx1)
    If myr >= 1999 And (mmn >= 1 And mmn <= 12) Then
        With rstTS
            .FindFirst "yr = " & myr & " and mn = " & mmn
            If Not .EOF Then
                If vbYes = MsgBox("Already processed. Process again  ", vbYesNo) Then
                    a = 1
                Else
                    a = 0
                End If
            Else
                a = 1
            End If
        End With
        If a = 1 Then
            With rstTS
                
                If .RecordCount > 0 Then
                While Not .EOF
                    .Delete
                    .MoveNext
                Wend
                End If
            End With
            
            With rstPS
                If .RecordCount > 0 Then
                .MoveFirst
                While Not .EOF
                    xDoIt = True
                    mjoin = False
                    mresign = False
                    mpfno = ""
                    mtotwk = 0
                    mwkdy = 0
                    mhldy = 0
                    mcldy = 0
                    meldy = 0
                    mabdy = 0
                    mtotwk = mtotdy
                    mcode = !EMPCODE
                    mbasic = !empbasic
                    mda = !empda
                    mhra = !emphra
                    mmed = !empmed
                    mca = !empcall
                    mlta = !empltall
                    mua = !empuall
                    mspl = !empsplall
                    
                    mabasic = mbasic
                    mada = mda
                    mahra = mhra
                    mamed = mmed
                    maca = mca
                    malta = mlta
                    maua = mua
                    maspl = mspl
                    Label5.Caption = "Processing Code " & mcode
                    With rstPM
                        .FindFirst "empcode = '" & mcode & "'"
                        If Not .EOF Then
                            mdesg = !empdesg
                            mdept = !empdept
                            mcomp = !empcomp
                            mxdoj = !EMPDOJ
                            mpfno = !emppf
                            If IsNull(!EMPDOR) Then
                            mxdor = CDate("01/01/1900")
                            Else
                            mxdor = !EMPDOR
                            End If
                            If Str(Year(mxdoj)) = Str(myr) And Str(Month(mxdoj)) = Str(mmn) Then
                                mjoin = True
                            Else
                                mjoin = False
                            End If
                            If Str(Year(mxdor)) = Str(myr) And Str(Month(mxdor)) = Str(mmn) Then
                                mresign = True
                                xDoIt = True
                            Else
                                If mxdor = CDate("01/01/1900") Then
                                    mresign = False
                                    xDoIt = True
                                Else
                                    xDoIt = False
                                End If
                            End If
                        Else
                            mdesg = " "
                            mdept = " "
                            mcomp = " "
                            mjoin = False
                            mresign = False
                            xDoIt = False
                        End If
                    End With
                    If xDoIt Then
                    mx = mxx
                    While mx <= mx1
                        If mjoin Then
                            If mx < mxdoj Then
                                mabdy = mabdy + 1
                            Else
                                If Weekday(mx) = 1 Then
                                    mhldy = mhldy + 1
                                End If
                            End If
                        Else
                            If mresign Then
                                If mx > mxdor Then
                                    mabdy = mabdy + 1
                                Else
                                    If Weekday(mx) = 1 Then
                                        mhldy = mhldy + 1
                                    End If
                                End If
                            Else
                                If Weekday(mx) = 1 Then
                                    mhldy = mhldy + 1
                                End If
                            End If
                        End If
                        mx = DateAdd("d", 1, mx)
                    Wend
                    With rstTL
                        .FindFirst "empcode = '" & mcode & "' and yr = " & myr & " and mn = " & mmn
                        intCtrTest = 0
                        While Not .NoMatch
                            intCtrTest = intCtrTest + 1
                            Label5.Caption = "Processing Code " & mcode & "  Leave Processing " & intCtrTest
                            If Weekday(!DOL) = 1 Then
                                mhldy = mhldy - 1
                              End If
                            Select Case !LEAVTYPE
                            Case "H"
                                mhldy = mhldy + !LEAVDAY
                            Case "C"
                                mcldy = mcldy + !LEAVDAY
                            Case "E"
                                meldy = meldy + !LEAVDAY
                            Case "A"
                                mabdy = mabdy + !LEAVDAY
                            Case "L"
                                mabdy = mabdy + !LEAVDAY
                            Case "P"
                                mPrdy = mPrdy + !LEAVDAY
                            End Select
                      
                            .FindNext "empcode = '" & mcode & "' and yr = " & myr & " and mn = " & mmn
                        Wend
                    End With
                    mtotwk = mtotwk - mabdy
                    mwkdy = mtotwk - mhldy - mcldy - meldy
                    If mtotdy <> mtotwk Then
                        mbasic = mbasic * (mtotwk / mtotdy)
                        mda = mda * (mtotwk / mtotdy)
                        mhra = mhra * (mtotwk / mtotdy)
                        mmed = mmed * (mtotwk / mtotdy)
                        mlta = mlta * (mtotwk / mtotdy)
                        mca = mca * (mtotwk / mtotdy)
                        mua = mua * (mtotwk / mtotdy)
                        mspl = mspl * (mtotwk / mtotdy)
                    End If
                    mpfamt = 0
                    mesiamt = 0
                    mpf = 0
                    mepf = 0
                    mesi = 0
                    xpf = 0
                    With rstOTS
                        .FindFirst "empcode = '" & Trim(mcode) & "'"
                        If Not .NoMatch Then
                            mLtESI = !empesi
                        Else
                            mLtESI = 0
                        End If
                    End With
                    
                    If mpfno <> "" Or (mabasic + mada) <= 6500 Then '5000
                        mbasic = Round(mbasic, 2)
                        If mbasic - Int(mbasic) >= 0.5 Then
                            mbasic = Int(mbasic) + 1
                        End If
                        mpfamt = mbasic + mda
                        xpf = ((mbasic + mda) * (12 / 100))
                        If xpf - Int(xpf) >= 0.5 Then
                            xpf = Int(xpf) + 1
                        Else
                            xpf = Int(xpf)
                        End If
                        mpf = (xpf * (3.67 / 100))
                        mpf = Int(mpf)
                        mepf = xpf - mpf
                        If (mepf + mpf) > 780 Then
                            mepf = 541 '417'
                            mpf = 239
                        End If
                    End If
                    If (mabasic + mada + mahra + maua + maspl) <= 6500 Then '5000
                        mbasic = Round(mbasic, 2)
                        If mbasic - Int(mbasic) >= 0.5 Then
                            mbasic = Int(mbasic) + 1
                        End If
                        mesiamt = mbasic + mda + mhra + mua + mspl
                        mesi = (mbasic + mda + mhra + mua + mspl + mca) * (1.75 / 100)
                        mesi = Round(mesi, 2)
                        xxctr = Val(Format$(mesi - Int(mesi), "#.00"))
                        If xxctr > 0 Then
                        xxctr = xxctr * 100
                        If (xxctr Mod 5) = 0 Then
                        Else
                        Do While (xxctr Mod 5) <> 0
                            xxctr = xxctr + 1
                        Loop
                        xxctr = (xxctr / 100)
                        mesi = Int(mesi) + xxctr
                        mesi = Round(mesi, 2)
                        End If
                        End If
                        'If Format$(Int(mesi), "#.00") <> Format$(mesi, "#.00") Then
                        '    mesi = mesi + 1
                        '    mesi = Int(mesi)
                        'End If
                    Else
                        If Val(txtMonth) = 10 Or Val(txtMonth) = 4 Then
                            mesi = 0
                        Else
                            If mLtESI > 0 Then
                            mesi = 113.75
                            Else
                            mesi = 0
                            End If
                        End If
                    End If
                    
                    With rstTS
                        .AddNew
                        !EMPCODE = mcode
                        !YR = myr
                        !MN = mmn
                        !empcomp = mcomp
                        !empdept = mdept
                        !empdesg = mdesg
                        !empbasic = mbasic
                        !empda = mda
                        !emphra = mhra
                        !empmed = mmed
                        !empcall = mca
                        !empltall = mlta
                        !empuall = mua
                        !empsplall = mspl
                        !actesiamt = mesiamt
                        !actpfamt = mpfamt
                        !empesi = mesi
                        !emppf = mpf
                        !empepf = mepf
                        !actbasic = mabasic
                        !actda = mada
                        !acthra = mahra
                        !actmed = mamed
                        !actcall = maca
                        !actltall = malta
                        !actuall = maua
                        !actsplall = maspl
                        !empwkdy = Format$(mwkdy, "#.00")
                        !emphldy = Format$(mhldy, "#.00")
                        !empcldy = Format$(mcldy, "#.00")
                        !empeldy = Format$(meldy, "#.00")
                        !empabdy = Format$(mabdy, "#.00")
                        .Update
                    End With
                    End If
                    .MoveNext
                Wend
                End If
            End With
        Else
        
        End If
    Else
        MsgBox "Invalid Month and Year "
        txtMonth.SetFocus
    End If
End If
If strVal = "2" Then
    CrystalReport1.WindowState = crptMaximized
'    CrystalReport1.ReportFileName = strPath & "salnewfor.rpt"
    CrystalReport1.ReportFileName = strPath & "salBank.rpt"
    CrystalReport1.SelectionFormula = "{transal.yr} = " & Val(Trim(txtYear)) & " and {transal.mn} = " & Val(Trim(txtMonth))
    CrystalReport1.Connect = "DSN = PMS"
'    CrystalReport1.Destination = crptToFile
'    CrystalReport1.PrinterStartPage = 1
'    CrystalReport1.PrinterStopPage = 1
    CrystalReport1.Action = 1
    
End If
End Sub

Private Sub cmdRemove_Click()
If SelList.ListIndex >= 0 Then
    SelList.RemoveItem (SelList.ListIndex)
End If
End Sub

Private Sub Command1_Click()
Dim intCtrTest As Integer
Dim mxx As Date
If txtMonth.Text = "" Then
    txtMonth.SetFocus
    Exit Sub
End If
If txtYear.Text = "" Then
    txtYear.SetFocus
    Exit Sub
End If
If strVal = "1" Then
    xOMN = Val(txtMonth) - 1
    If xOMN = 0 Then
        xOMN = 12
        xOYR = Val(txtYear) - 1
    Else
        xOYR = Val(txtYear)
    End If
    Set dbpms = OpenDatabase(strPath & "dbpms.mdb")
    Set rstPM = dbpms.OpenRecordset("select * from mastpay", dbOpenDynaset)
    Set rstPS = dbpms.OpenRecordset("select * from mastsal order by empcode", dbOpenDynaset)
    Set rstTS = dbpms.OpenRecordset("select * from TRANSAL where yr = " & Val(txtYear.Text) & " and mn = " & Val(txtMonth))
    Set rstTL = dbpms.OpenRecordset("select * from TRANLEAV", dbOpenDynaset)
    Set rstOTS = dbpms.OpenRecordset("select * from TRANSAL where yr = " & xOYR & " and mn = " & xOMN)
    Set rstMNBL = dbpms.OpenRecordset("select * from mnblock", dbOpenDynaset)
    myr = txtYear
    mmn = txtMonth
    
    With rstMNBL
      .MoveFirst
      Do While Not .EOF
         If !yrblock = myr And !mnblock = mmn Then
            MsgBox "Unable to Calculate. Month Blocked "
            Exit Sub
         End If
         .MoveNext
      Loop
    End With
    mx = txtMonth & "/02/" & txtYear '01
    mx1 = DateAdd("m", 1, mx)
    mx1 = DateAdd("d", -1, mx1)
    mxx = mx
    mtotdy = Day(mx1)
    If myr >= 1999 And (mmn >= 1 And mmn <= 12) Then
        With rstTS
            .FindFirst "yr = " & myr & " and mn = " & mmn
            If Not .EOF Then
                If vbYes = MsgBox("Already processed. Process again  ", vbYesNo) Then
                    a = 1
                Else
                    a = 0
                End If
            Else
                a = 1
            End If
        End With
        If a = 1 Then
            With rstTS
                
                If .RecordCount > 0 Then
                While Not .EOF
                    .Delete
                    .MoveNext
                Wend
                End If
            End With
            
            With rstPS
                If .RecordCount > 0 Then
                .MoveFirst
                While Not .EOF
                    xDoIt = True
                    mjoin = False
                    mresign = False
                    mpfno = ""
                    mtotwk = 0
                    mwkdy = 0
                    mhldy = 0
                    mcldy = 0
                    meldy = 0
                    mabdy = 0
                    mtotwk = mtotdy
                    mcode = !EMPCODE
                    mbasic = !empbasic
                    mda = !empda
                    mhra = !emphra
                    mmed = !empmed
                    mca = !empcall
                    mlta = !empltall
                    mua = !empuall
                    mspl = !empsplall
                    
                    mabasic = mbasic
                    mada = mda
                    mahra = mhra
                    mamed = mmed
                    maca = mca
                    malta = mlta
                    maua = mua
                    maspl = mspl
                    Label5.Caption = "Processing Code " & mcode
                    With rstPM
                        .FindFirst "empcode = '" & mcode & "'"
                        If Not .EOF Then
                            mdesg = !empdesg
                            mdept = !empdept
                            mcomp = !empcomp
                            mxdoj = !EMPDOJ
                            mpfno = !emppf
                            If IsNull(!EMPDOR) Then
                            mxdor = CDate("01/01/2000")
                            Else
                            mxdor = !EMPDOR
                            End If
                            If Str(Year(mxdoj)) = Str(myr) And Str(Month(mxdoj)) = Str(mmn) Then
                                mjoin = True
                            Else
                                mjoin = False
                            End If
                            If Str(Year(mxdor)) = Str(myr) And Str(Month(mxdor)) = Str(mmn) Then
                                mresign = True
                                xDoIt = True
                            Else
                                If mxdor = CDate("01/01/2000") Then
                                    mresign = False
                                    xDoIt = True
                                Else
                                    xDoIt = False
                                End If
                            End If
                        Else
                            mdesg = " "
                            mdept = " "
                            mcomp = " "
                            mjoin = False
                            mresign = False
                            xDoIt = False
                        End If
                    End With
                    If xDoIt Then
                    mx = mxx
                    While mx <= mx1
                        If mjoin Then
                            If mx < mxdoj Then
                                mabdy = mabdy + 1
                            Else
                                If Weekday(mx) = 1 Then
                                    mhldy = mhldy + 1
                                End If
                            End If
                        Else
                            If mresign Then
                                If mx > mxdor Then
                                    mabdy = mabdy + 1
                                Else
                                    If Weekday(mx) = 1 Then
                                        mhldy = mhldy + 1
                                    End If
                                End If
                            Else
                                If Weekday(mx) = 1 Then
                                    mhldy = mhldy + 1
                                End If
                            End If
                        End If
                        mx = DateAdd("d", 1, mx)
                    Wend
                    With rstTL
                        .FindFirst "empcode = '" & mcode & "' and yr = " & myr & " and mn = " & mmn
                        intCtrTest = 0
                        While Not .NoMatch
                            intCtrTest = intCtrTest + 1
                            Label5.Caption = "Processing Code " & mcode & "  Leave Processing " & intCtrTest
                            If Weekday(!DOL) = 1 Then
                                mhldy = mhldy - 1
                              End If
                            Select Case !LEAVTYPE
                            Case "H"
                                mhldy = mhldy + !LEAVDAY
                            Case "C"
                                mcldy = mcldy + !LEAVDAY
                            Case "E"
                                meldy = meldy + !LEAVDAY
                            Case "A"
                                mabdy = mabdy + !LEAVDAY
                            Case "L"
                                mabdy = mabdy + !LEAVDAY
                            Case "P"
                                mPrdy = mPrdy + !LEAVDAY
                            End Select
                      
                            .FindNext "empcode = '" & mcode & "' and yr = " & myr & " and mn = " & mmn
                        Wend
                    End With
                    mtotwk = mtotwk - mabdy
                    mwkdy = mtotwk - mhldy - mcldy - meldy
                    If mtotdy <> mtotwk Then
                        mbasic = mbasic * (mtotwk / mtotdy)
                        mda = mda * (mtotwk / mtotdy)
                        mhra = mhra * (mtotwk / mtotdy)
                        mmed = mmed * (mtotwk / mtotdy)
                        mlta = mlta * (mtotwk / mtotdy)
                        mca = mca * (mtotwk / mtotdy)
                        mua = mua * (mtotwk / mtotdy)
                        mspl = mspl * (mtotwk / mtotdy)
                    End If
                    mpfamt = 0
                    mesiamt = 0
                    mpf = 0
                    mepf = 0
                    mesi = 0
                    xpf = 0
                    With rstOTS
                        .FindFirst "empcode = '" & Trim(mcode) & "'"
                        If Not .NoMatch Then
                            mLtESI = !empesi
                        Else
                            mLtESI = 0
                        End If
                    End With
                    
                    If mpfno <> "" Or (mabasic + mada) <= 6500 Then ' CHANGED FROM 5000
                        mbasic = Round(mbasic, 2)
                        If mbasic - Int(mbasic) >= 0.5 Then
                            mbasic = Int(mbasic) + 1
                        End If
                        mpfamt = mbasic + mda
                        xpf = ((mbasic + mda) * (12 / 100))
                        If xpf - Int(xpf) >= 0.5 Then
                            xpf = Int(xpf) + 1
                        Else
                            xpf = Int(xpf)
                        End If
                        mpf = (xpf * (3.67 / 100))
                        mpf = Int(mpf)
                        mepf = xpf - mpf
                        If (mepf + mpf) > 780 Then
                            mepf = 541 '417'
                            mpf = 239
                        End If
                    End If
                    If (mabasic + mada + mahra + maua + maspl) <= 6500 Then
                        mbasic = Round(mbasic, 2)
                        If mbasic - Int(mbasic) >= 0.5 Then
                            mbasic = Int(mbasic) + 1
                        End If
                        mesiamt = mbasic + mda + mhra + mua + mspl
                        mesi = (mbasic + mda + mhra + mua + mspl) * (1.75 / 100)
                        mesi = Round(mesi, 2)
                        xxctr = Val(Format$(mesi - Int(mesi), "#.00"))
                        If xxctr > 0 Then
                        xxctr = xxctr * 100
                        If (xxctr Mod 5) = 0 Then
                        Else
                        Do While (xxctr Mod 5) <> 0
                            xxctr = xxctr + 1
                        Loop
                        xxctr = (xxctr / 100)
                        mesi = Int(mesi) + xxctr
                        mesi = Round(mesi, 2)
                        End If
                        End If
                        'If Format$(Int(mesi), "#.00") <> Format$(mesi, "#.00") Then
                        '    mesi = mesi + 1
                        '    mesi = Int(mesi)
                        'End If
                    Else
                        If Val(txtMonth) = 10 Or Val(txtMonth) = 4 Then
                            mesi = 0
                        Else
                            If mLtESI > 0 Then
                            mesi = 113.75
                            Else
                            mesi = 0
                            End If
                        End If
                    End If
                    
                    With rstTS
                        .AddNew
                        !EMPCODE = mcode
                        !YR = myr
                        !MN = mmn
                        !empcomp = mcomp
                        !empdept = mdept
                        !empdesg = mdesg
                        !empbasic = mbasic
                        !empda = mda
                        !emphra = mhra
                        !empmed = mmed
                        !empcall = mca
                        !empltall = mlta
                        !empuall = mua
                        !empsplall = mspl
                        !actesiamt = mesiamt
                        !actpfamt = mpfamt
                        !empesi = mesi
                        !emppf = mpf
                        !empepf = mepf
                        !actbasic = mabasic
                        !actda = mada
                        !acthra = mahra
                        !actmed = mamed
                        !actcall = maca
                        !actltall = malta
                        !actuall = maua
                        !actsplall = maspl
                        !empwkdy = Format$(mwkdy, "#.00")
                        !emphldy = Format$(mhldy, "#.00")
                        !empcldy = Format$(mcldy, "#.00")
                        !empeldy = Format$(meldy, "#.00")
                        !empabdy = Format$(mabdy, "#.00")
                        .Update
                    End With
                    End If
                    .MoveNext
                Wend
                End If
            End With
        Else
        
        End If
    Else
        MsgBox "Invalid Month and Year "
        txtMonth.SetFocus
    End If
End If
If strVal = "2" Then
    CrystalReport1.WindowState = crptMaximized
'    CrystalReport1.ReportFileName = strPath & "salnewfor.rpt"
    CrystalReport1.ReportFileName = strPath & "MANAG.rpt"
    CrystalReport1.SelectionFormula = "{transal.yr} = " & Val(Trim(txtYear)) & " and {transal.mn} = " & Val(Trim(txtMonth))
    CrystalReport1.Connect = "DSN = PMS"
'    CrystalReport1.Destination = crptToFile
'    CrystalReport1.PrinterStartPage = 1
'    CrystalReport1.PrinterStopPage = 1
    CrystalReport1.Action = 1
    
End If
End Sub

Private Sub cmdresign_Click()
Dim intCtrTest As Integer
Dim mxx As Date
If txtMonth.Text = "" Then
    txtMonth.SetFocus
    Exit Sub
End If
If txtYear.Text = "" Then
    txtYear.SetFocus
    Exit Sub
End If
If strVal = "1" Then
    xOMN = Val(txtMonth) - 1
    If xOMN = 0 Then
        xOMN = 12
        xOYR = Val(txtYear) - 1
    Else
        xOYR = Val(txtYear)
    End If
    Set dbpms = OpenDatabase(strPath & "dbpms.mdb")
    Set rstPM = dbpms.OpenRecordset("select * from mastpay", dbOpenDynaset)
    Set rstPS = dbpms.OpenRecordset("select * from mastsal order by empcode", dbOpenDynaset)
    Set rstTS = dbpms.OpenRecordset("select * from TRANSAL where yr = " & Val(txtYear.Text) & " and mn = " & Val(txtMonth))
    Set rstTL = dbpms.OpenRecordset("select * from TRANLEAV", dbOpenDynaset)
    Set rstOTS = dbpms.OpenRecordset("select * from TRANSAL where yr = " & xOYR & " and mn = " & xOMN)
    Set rstMNBL = dbpms.OpenRecordset("select * from mnblock", dbOpenDynaset)
    myr = txtYear
    mmn = txtMonth
    
    With rstMNBL
      .MoveFirst
      Do While Not .EOF
         If !yrblock = myr And !mnblock = mmn Then
            MsgBox "Unable to Calculate. Month Blocked "
            Exit Sub
         End If
         .MoveNext
      Loop
    End With
    mx = txtMonth & "/01/" & txtYear
    mx1 = DateAdd("m", 1, mx)
    mx1 = DateAdd("d", -1, mx1)
    mxx = mx
    mtotdy = Day(mx1)
    If myr >= 1999 And (mmn >= 1 And mmn <= 12) Then
        With rstTS
            .FindFirst "yr = " & myr & " and mn = " & mmn
            If Not .EOF Then
                If vbYes = MsgBox("Already processed. Process again  ", vbYesNo) Then
                    a = 1
                Else
                    a = 0
                End If
            Else
                a = 1
            End If
        End With
        If a = 1 Then
            With rstTS
                
                If .RecordCount > 0 Then
                While Not .EOF
                    .Delete
                    .MoveNext
                Wend
                End If
            End With
            
            With rstPS
                If .RecordCount > 0 Then
                .MoveFirst
                While Not .EOF
                    xDoIt = True
                    mjoin = False
                    mresign = False
                    mpfno = ""
                    mtotwk = 0
                    mwkdy = 0
                    mhldy = 0
                    mcldy = 0
                    meldy = 0
                    mabdy = 0
                    mtotwk = mtotdy
                    mcode = !EMPCODE
                    mbasic = !empbasic
                    mda = !empda
                    mhra = !emphra
                    mmed = !empmed
                    mca = !empcall
                    mlta = !empltall
                    mua = !empuall
                    mspl = !empsplall
                    
                    mabasic = mbasic
                    mada = mda
                    mahra = mhra
                    mamed = mmed
                    maca = mca
                    malta = mlta
                    maua = mua
                    maspl = mspl
                    Label5.Caption = "Processing Code " & mcode
                    With rstPM
                        .FindFirst "empcode = '" & mcode & "'"
                        If Not .EOF Then
                            mdesg = !empdesg
                            mdept = !empdept
                            mcomp = !empcomp
                            mxdoj = !EMPDOJ
                            mpfno = !emppf
                            If IsNull(!EMPDOR) Then
                            mxdor = CDate("01/01/2000")
                            Else
                            mxdor = !EMPDOR
                            End If
                            If Str(Year(mxdoj)) = Str(myr) And Str(Month(mxdoj)) = Str(mmn) Then
                                mjoin = True
                            Else
                                mjoin = False
                            End If
                            If Str(Year(mxdor)) = Str(myr) And Str(Month(mxdor)) = Str(mmn) Then
                                mresign = True
                                xDoIt = True
                            Else
                                If mxdor = CDate("01/01/2000") Then
                                    mresign = False
                                    xDoIt = True
                                Else
                                    xDoIt = False
                                End If
                            End If
                        Else
                            mdesg = " "
                            mdept = " "
                            mcomp = " "
                            mjoin = False
                            mresign = False
                            xDoIt = False
                        End If
                    End With
                    If xDoIt Then
                    mx = mxx
                    While mx <= mx1
                        If mjoin Then
                            If mx < mxdoj Then
                                mabdy = mabdy + 1
                            Else
                                If Weekday(mx) = 1 Then
                                    mhldy = mhldy + 1
                                End If
                            End If
                        Else
                            If mresign Then
                                If mx > mxdor Then
                                    mabdy = mabdy + 1
                                Else
                                    If Weekday(mx) = 1 Then
                                        mhldy = mhldy + 1
                                    End If
                                End If
                            Else
                                If Weekday(mx) = 1 Then
                                    mhldy = mhldy + 1
                                End If
                            End If
                        End If
                        mx = DateAdd("d", 1, mx)
                    Wend
                    With rstTL
                        .FindFirst "empcode = '" & mcode & "' and yr = " & myr & " and mn = " & mmn
                        intCtrTest = 0
                        While Not .NoMatch
                            intCtrTest = intCtrTest + 1
                            Label5.Caption = "Processing Code " & mcode & "  Leave Processing " & intCtrTest
                            If Weekday(!DOL) = 1 Then
                                mhldy = mhldy - 1
                              End If
                            Select Case !LEAVTYPE
                            Case "H"
                                mhldy = mhldy + !LEAVDAY
                            Case "C"
                                mcldy = mcldy + !LEAVDAY
                            Case "E"
                                meldy = meldy + !LEAVDAY
                            Case "A"
                                mabdy = mabdy + !LEAVDAY
                            Case "L"
                                mabdy = mabdy + !LEAVDAY
                            Case "P"
                                mPrdy = mPrdy + !LEAVDAY
                            End Select
                      
                            .FindNext "empcode = '" & mcode & "' and yr = " & myr & " and mn = " & mmn
                        Wend
                    End With
                    mtotwk = mtotwk - mabdy
                    mwkdy = mtotwk - mhldy - mcldy - meldy
                    If mtotdy <> mtotwk Then
                        mbasic = mbasic * (mtotwk / mtotdy)
                        mda = mda * (mtotwk / mtotdy)
                        mhra = mhra * (mtotwk / mtotdy)
                        mmed = mmed * (mtotwk / mtotdy)
                        mlta = mlta * (mtotwk / mtotdy)
                        mca = mca * (mtotwk / mtotdy)
                        mua = mua * (mtotwk / mtotdy)
                        mspl = mspl * (mtotwk / mtotdy)
                    End If
                    mpfamt = 0
                    mesiamt = 0
                    mpf = 0
                    mepf = 0
                    mesi = 0
                    xpf = 0
                    With rstOTS
                        .FindFirst "empcode = '" & Trim(mcode) & "'"
                        If Not .NoMatch Then
                            mLtESI = !empesi
                        Else
                            mLtESI = 0
                        End If
                    End With
                    
                    If mpfno <> "" Or (mabasic + mada) <= 5000 Then
                        mbasic = Round(mbasic, 2)
                        If mbasic - Int(mbasic) >= 0.5 Then
                            mbasic = Int(mbasic) + 1
                        End If
                        mpfamt = mbasic + mda
                        xpf = ((mbasic + mda) * (12 / 100))
                        If xpf - Int(xpf) >= 0.5 Then
                            xpf = Int(xpf) + 1
                        Else
                            xpf = Int(xpf)
                        End If
                        mpf = (xpf * (3.67 / 100))
                        mpf = Int(mpf)
                        mepf = xpf - mpf
                        If (mepf + mpf) > 780 Then
                            mepf = 541 '417'
                            mpf = 239
                        End If
                    End If
                    If (mabasic + mada + mahra + maua + maspl) <= 6500 Then
                        mbasic = Round(mbasic, 2)
                        If mbasic - Int(mbasic) >= 0.5 Then
                            mbasic = Int(mbasic) + 1
                        End If
                        mesiamt = mbasic + mda + mhra + mua + mspl
                        mesi = (mbasic + mda + mhra + mua + mspl) * (1.75 / 100)
                        mesi = Round(mesi, 2)
                        xxctr = Val(Format$(mesi - Int(mesi), "#.00"))
                        If xxctr > 0 Then
                        xxctr = xxctr * 100
                        If (xxctr Mod 5) = 0 Then
                        Else
                        Do While (xxctr Mod 5) <> 0
                            xxctr = xxctr + 1
                        Loop
                        xxctr = (xxctr / 100)
                        mesi = Int(mesi) + xxctr
                        mesi = Round(mesi, 2)
                        End If
                        End If
                        'If Format$(Int(mesi), "#.00") <> Format$(mesi, "#.00") Then
                        '    mesi = mesi + 1
                        '    mesi = Int(mesi)
                        'End If
                    Else
                        If Val(txtMonth) = 10 Or Val(txtMonth) = 4 Then
                            mesi = 0
                        Else
                            If mLtESI > 0 Then
                            mesi = 113.75
                            Else
                            mesi = 0
                            End If
                        End If
                    End If
                    
                    With rstTS
                        .AddNew
                        !EMPCODE = mcode
                        !YR = myr
                        !MN = mmn
                        !empcomp = mcomp
                        !empdept = mdept
                        !empdesg = mdesg
                        !empbasic = mbasic
                        !empda = mda
                        !emphra = mhra
                        !empmed = mmed
                        !empcall = mca
                        !empltall = mlta
                        !empuall = mua
                        !empsplall = mspl
                        !actesiamt = mesiamt
                        !actpfamt = mpfamt
                        !empesi = mesi
                        !emppf = mpf
                        !empepf = mepf
                        !actbasic = mabasic
                        !actda = mada
                        !acthra = mahra
                        !actmed = mamed
                        !actcall = maca
                        !actltall = malta
                        !actuall = maua
                        !actsplall = maspl
                        !empwkdy = Format$(mwkdy, "#.00")
                        !emphldy = Format$(mhldy, "#.00")
                        !empcldy = Format$(mcldy, "#.00")
                        !empeldy = Format$(meldy, "#.00")
                        !empabdy = Format$(mabdy, "#.00")
                        .Update
                    End With
                    End If
                    .MoveNext
                Wend
                End If
            End With
        Else
        
        End If
    Else
        MsgBox "Invalid Month and Year "
        txtMonth.SetFocus
    End If
End If
If strVal = "2" Then
    CrystalReport1.WindowState = crptMaximized
'    CrystalReport1.ReportFileName = strPath & "salnewfor.rpt"
    CrystalReport1.ReportFileName = strPath & "salRegn.rpt"
    CrystalReport1.SelectionFormula = "{transal.yr} = " & Val(Trim(txtYear)) & " and {transal.mn} = " & Val(Trim(txtMonth))
    CrystalReport1.Connect = "DSN = PMS"
'    CrystalReport1.Destination = crptToFile
'    CrystalReport1.PrinterStartPage = 1
'    CrystalReport1.PrinterStopPage = 1
    CrystalReport1.Action = 1
    
End If
End Sub

Private Sub Command2_Click()
    frmMain.Enabled = True
    Unload frmCalSal
    frmMain.Show
End Sub

Private Sub Form_Load()
If strVal = "2" Then
    frmCalSal.Caption = "Salary Register"
    TotList.AddItem ("Employee Code")
    TotList.AddItem ("Department")
    TotList.AddItem ("Designation")
    TotList.ListIndex = 0
    cmdAdd_Click
Else
    frmCalSal.Caption = "Salary Calculation"
    TotList.Enabled = False
    SelList.Enabled = False
    frmCalSal.cmdRemove.Enabled = False
    frmCalSal.cmdAdd.Enabled = False
End If

End Sub

Private Sub MSHFlexGrid1_Click()

End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
    End If
End Sub
