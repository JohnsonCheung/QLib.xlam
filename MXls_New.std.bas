Attribute VB_Name = "MXls_New"
Option Explicit
Function NewA1(Optional Wsn$, Optional Vis As Boolean) As Range
Set NewA1 = RgVis(A1zWs(NewWs(Wsn)))
End Function
Function NewWb(Optional Wsn$) As Workbook
Dim O As Workbook
Dim X As Excel.Application
Set O = NewXls.Workbooks.Add
Set NewWb = WbzWs(SetWsNm(FstWs(O), Wsn))
End Function

Function NewWs(Optional Wsn$) As Worksheet
Set NewWs = SetWsNm(FstWs(NewWb), Wsn)
End Function

Function NewXls() As Excel.Application
Set NewXls = CreateObject("Excel.Application") ' Don't use New Excel.Application
End Function
