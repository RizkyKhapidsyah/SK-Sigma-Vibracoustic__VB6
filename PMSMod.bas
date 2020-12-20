Attribute VB_Name = "PMSMod"
Option Explicit
Public strVal, strVal1  As String
Global strPath As String
Sub main()
strPath = "c:\pms\dbpms\"
Load frmMain
Load frmSplash
frmMain.Show
frmSplash.Show
End Sub

Public Function mnuViewOrder() As String
Dim i As Integer
Dim strVal As String
For i = 0 To 5 Step 1
If frmMain.mnuViewName(i).Checked = True Then
strVal = "select empcode, empname, empdob, empdoj, empesi, emppf, mastdesg.desgdes from mastpay, mastdesg where mastdesg.empdesg = mastpay.empdesg order by "
Select Case i
Case 0:
strVal = strVal & "mastpay.empName"
Case 1:
strVal = strVal & "mastpay.empCode"
Case 2:
strVal = strVal & "mastpay.empesi"
Case 3:
strVal = strVal & "mastpay.emppf"
Case 4:
strVal = strVal & "mastpay.empdoj"
Case 5:
strVal = strVal & "mastpay.empdob"
End Select
End If
Next i
mnuViewOrder = strVal
End Function

