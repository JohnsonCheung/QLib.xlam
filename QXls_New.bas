Attribute VB_Name = "QXls_New"
Option Explicit
Private Const CMod$ = "MXls_New."
Private Const Asm$ = "QXls"
Function NewA1(Optional Wsn$, Optional Vis As Boolean) As Range
Set NewA1 = ShwRg(A1zWs(NewWs(Wsn)))
End Function

Function NewWb(Optional Wsn$) As Workbook
Dim O As Workbook
Set O = Xls.Workbooks.Add
Set NewWb = WbzWs(SetWsn(FstWs(O), Wsn))
End Function

Function NewWs(Optional Wsn$) As Worksheet
Set NewWs = SetWsn(FstWs(NewWb), Wsn)
End Function

Function NewXls() As Excel.Application
Set NewXls = CreateObject("Excel.Application") ' Don't use New Excel.Application
End Function
