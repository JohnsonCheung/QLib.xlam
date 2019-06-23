Attribute VB_Name = "QXls_AddIn"
Option Compare Text
Option Explicit
Private Const CMod$ = "MXls_AddIn."
Private Const Asm$ = "QXls"

Function DAddin(A As Excel.Application) As Drs
DAddin = DrszItrPP(A.AddIns, "Name FullName Installed IsOpen ProgId CLSID")
End Function

Sub DmpAddinX()
DmpAddinzX Xls
End Sub

Sub DmpAddinzX(X As Excel.Application)
DmpDrs DAddin(X)
End Sub

Function WsAddin(X As Excel.Application) As Worksheet
Set WsAddin = ShwWs(WszDrs(DAddin(X)))
End Function

Function Addin(X As Excel.Application, FxaNm) As Excel.Addin
Dim I As Excel.Addin
For Each I In X.AddIns
    If I.Name = FxaNm & ".xlam" Then Set Addin = I
Next
End Function

