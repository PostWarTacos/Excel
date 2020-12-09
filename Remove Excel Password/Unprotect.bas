Attribute VB_Name = "Unprotect"
Sub UnprotectAllSheets()
Attribute UnprotectAllSheets.VB_ProcData.VB_Invoke_Func = "U\n14"
'Shortcut: Ctrl + Shift + U

Dim ws As Worksheet
Dim wb As Workbook
Dim strPwd As String
Dim strCheck As String
Set wb = ActiveWorkbook

strCheck = "DCSOPS"
strPwd = InputBox("Enter Password", "Password", "Enter Password")

If strPwd = strCheck Then
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect Password:=strPwd
    Next ws
        wb.Unprotect Password:="DCSOPS"
Else
    MsgBox "Incorrect Password"
End If

End Sub
