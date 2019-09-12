Attribute VB_Name = "MxLoPrp"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxLoPrp."
Function LoColRg(L As ListObject, F) As Range
Set LoColRg = L.ListColumns(F).DataBodyRange.EntireColumn
End Function


