Attribute VB_Name = "QXls_AddIn"
Option Compare Text
Option Explicit
Private Const CMod$ = "MXls_AddIn."
Private Const Asm$ = "QXls"

Function DAddins(A As Excel.Application) As Drs
'DAddins = DrszItrPP(A.AddIns, "Name FullName Installed IsOpen ProgId CLSID")
End Function
Sub DmpAddinsXls()
DmpAddins Xls
End Sub

Sub DmpAddins(A As Excel.Application)
DmpDrs DAddins(A)
End Sub

Function AddinsWs(A As Excel.Application) As Worksheet
Set AddinsWs = ShwWs(WszDrs(DAddins(A)))
End Function

Function Addin(A As Excel.Application, FxaNm) As Excel.Addin
Dim I As Excel.Addin
For Each I In A.AddIns
    If I.Name = FxaNm & ".xlam" Then Set Addin = I
Next
End Function

