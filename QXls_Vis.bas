Attribute VB_Name = "QXls_Vis"
Option Explicit
Private Const CMod$ = "MXls_Vis."
Private Const Asm$ = "QXls"

Function ShwWb(A As Workbook) As Workbook
ShwXls A.Application
Set ShwWb = A
End Function

Function ShwXls(A As Excel.Application) As Excel.Application
If Not A.Visible Then A.Visible = True
Set ShwXls = A
End Function

Function ShwRg(A As Range) As Range
ShwXls A.Application
Set ShwRg = A
End Function

Function ShwLo(A As ListObject) As ListObject
ShwXls A.Application
Set ShwLo = A
End Function

Function ShwWs(A As Worksheet) As Worksheet
ShwXls A.Application
Set ShwWs = A
End Function

