Attribute VB_Name = "QXls_Xls"
Option Explicit
Private Const CMod$ = "MXls_Xls."
Private Const Asm$ = "QXls"
Public Const XlsPgmFfn$ = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
Private Sub Z_XlsOfGetObj()
Debug.Print XlsOfGetObj.Name
End Sub

Function XlsOfGetObj() As Excel.Application
Set XlsOfGetObj = GetObject(XlsPgmFfn)
End Function

Function Xls() As Excel.Application
Set Xls = Excel.Application
End Function

Function HasAddinFn(A As Excel.Application, AddinFn$) As Boolean
HasAddinFn = HasItn(A.AddIns, AddinFn)
End Function

Sub QuitXls(A As Excel.Application)
Stamp "QuitXls: Start"
Stamp "QuitXls: ClsAllWb":    ClsAllWb A
Stamp "QuitXls: Quit":        A.Quit
Stamp "QuitXls: Set nothing": Set A = Nothing
Stamp "QuitXls: Done"
End Sub
Sub ClsAllWb(A As Excel.Application)
Dim W As Workbook
For Each W In A.Workbooks
    W.Close False
Next
End Sub

Function DftXls(A As Excel.Application) As Excel.Application
If IsNothing(A) Then
    Set DftXls = NewXls
Else
    Set DftXls = A
End If
End Function


