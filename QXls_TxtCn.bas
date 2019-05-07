Attribute VB_Name = "QXls_TxtCn"
Option Explicit
Private Const CMod$ = "MXls_TxtCn."
Private Const Asm$ = "QXls"
Function TxtCnzWc(A As WorkbookConnection) As TextConnection
On Error Resume Next
Set TxtCnzWc = A.TextConnection
End Function
