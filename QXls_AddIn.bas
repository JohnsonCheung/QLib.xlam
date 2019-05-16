Attribute VB_Name = "QXls_AddIn"
Option Explicit
Private Const CMod$ = "MXls_AddIn."
Private Const Asm$ = "QXls"

Function Drs_Addins(A As Excel.Application) As Drs
'Drs_Addins = DrszItrPP(A.AddIns, "Name FullName Installed IsOpen ProgId CLSID")
End Function
Sub DmpAddinsXls()
DmpAddins Xls
End Sub

Sub DmpAddins(A As Excel.Application)
DmpDrs Drs_Addins(A)
End Sub

Function AddinsWs(A As Excel.Application) As Worksheet
Set AddinsWs = ShwWs(WszDrs(Drs_Addins(A)))
End Function

Function Addin(A As Excel.Application, FxaNm) As Excel.Addin
Dim I As Excel.Addin
For Each I In A.AddIns
    If I.Name = FxaNm & ".xlam" Then Set Addin = I
Next
End Function

