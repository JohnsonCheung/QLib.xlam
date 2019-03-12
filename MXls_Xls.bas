Attribute VB_Name = "MXls_Xls"
Option Explicit
Function Xls() As Excel.Application
Set Xls = Excel.Application
End Function

Function HasAddInFn(A As Excel.Application, AddInFn) As Boolean
HasAddInFn = HasItn(A.AddIns, AddInFn)
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


