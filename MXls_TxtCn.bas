Attribute VB_Name = "MXls_TxtCn"
Option Explicit
Function TxtCnzWc(A As WorkbookConnection) As TextConnection
On Error Resume Next
Set TxtCnzWc = A.TextConnection
End Function
