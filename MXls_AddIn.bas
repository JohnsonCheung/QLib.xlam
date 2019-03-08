Attribute VB_Name = "MXls_AddIn"
Option Explicit

Function AddInsDrs(A As Excel.Application) As Drs
Set AddInsDrs = DrszItrPP(A.AddIns, "Name FullName Installed IsOpen ProgId CLSID")
End Function
Sub DmpAddInsXls()
DmpAddIns Xls
End Sub

Sub DmpAddIns(A As Excel.Application)
DmpDrs AddInsDrs(A)
End Sub

Function AddInsWs(A As Excel.Application) As Worksheet
Set AddInsWs = WsVis(WszDrs(AddInsDrs(A)))
End Function

Function AddIn(A As Excel.Application, FxaNm) As Excel.AddIn
Dim I As Excel.AddIn
For Each I In A.AddIns
    If I.Name = FxaNm & ".xlam" Then Set AddIn = I
Next
End Function

