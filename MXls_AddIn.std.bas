Attribute VB_Name = "MXls_AddIn"
Option Explicit

Function XlsAddInDrs(A As Excel.Application) As Drs
'Set XlsAddInDrs = ItrDrs(A.AddIns, "Name FullName Installed IsOpen ProgId CLSID")
End Function

Sub XlsAddInDmp(A As Excel.Application)
DmpDrs XlsAddInDrs(A)
End Sub

Property Get XlsAddInWs() As Worksheet
'Set XlsAddInWs = WsVis(WszDrs(XlsAddInDrs(Xls)))
End Property

