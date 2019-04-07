Attribute VB_Name = "MXls_Xls"
Option Explicit
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

Function HasAddinFn(A As Excel.Application, AddinFn) As Boolean
HasAddinFn = HasItn(A.AddIns, AddinFn)
End Function

Sub XlsQuit(A As Excel.Application)
Stamp "XlsQuit: Start"
Stamp "XlsQuit: ClsAllWb":    ClsAllWb A
Stamp "XlsQuit: Quit":        A.Quit
Stamp "XlsQuit: Set nothing": Set A = Nothing
Stamp "XlsQuit: Done"
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


