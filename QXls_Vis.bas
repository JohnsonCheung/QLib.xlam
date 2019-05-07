Attribute VB_Name = "QXls_Vis"
Option Explicit
Private Const CMod$ = "MXls_Vis."
Private Const Asm$ = "QXls"

Sub ShwWb(A As Workbook)
ShwXls A.Application
End Sub
Sub ShwXls(A As Excel.Application)
If Not A.Visible Then A.Visible = True
End Sub

Sub ShwRg(A As Range)
ShwXls A.Application
End Sub

Sub ShwLo(A As ListObject)
ShwXls A.Application
End Sub

