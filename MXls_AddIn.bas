Attribute VB_Name = "MXls_AddIn"
Option Explicit

Function AddinsDrs(A As Excel.Application) As Drs
Set AddinsDrs = DrszItrpp(A.AddIns, "Name FullName Installed IsOpen ProgId CLSID")
End Function
Sub DmpAddinsXls()
DmpAddins Xls
End Sub

Sub DmpAddins(A As Excel.Application)
DmpDrs AddinsDrs(A)
End Sub

Function AddinsWs(A As Excel.Application) As Worksheet
Set AddinsWs = WsVis(WszDrs(AddinsDrs(A)))
End Function

Function Addin(A As Excel.Application, FxaNm) As Excel.Addin
Dim I As Excel.Addin
For Each I In A.AddIns
    If I.Name = FxaNm & ".xlam" Then Set Addin = I
Next
End Function

