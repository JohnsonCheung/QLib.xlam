Attribute VB_Name = "MXls_Lo_Get_Lo"
Option Explicit

Function LoAy(A As Workbook) As ListObject()
Dim Ws As Worksheet
For Each Ws In A.Sheets
    PushObjzItr LoAy, Ws.ListObjects
Next
End Function

Function LozWs(A As Worksheet, LoNm$) As ListObject 'Return LoOpt
Set LozWs = FstItmzNm(A.ListObjects, LoNm)
End Function

Function FstLo(A As Worksheet) As ListObject 'Return LoOpt
Set FstLo = FstItm(A.ListObjects)
End Function

