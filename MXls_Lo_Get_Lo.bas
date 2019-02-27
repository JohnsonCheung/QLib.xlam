Attribute VB_Name = "MXls_Lo_Get_Lo"
Option Explicit
Function LozWb(A As Workbook, LoNm$) As ListObject
Dim Ws As Worksheet
For Each Ws In A.Sheets
    If HasLo(Ws, LoNm) Then Set LozWb = Ws.ListObjects(LoNm): Exit Function
Next
End Function

Function LoAy(A As Workbook) As ListObject()
Dim Ws As Worksheet
For Each Ws In A.Sheets
    PushObjItr LoAy, Ws.ListObjects
Next
End Function

Function LozWs(A As Worksheet, LoNm$) As ListObject
Set LozWs = FstItrNm(A.ListObjects, LoNm)
End Function

Function FstLo(A As Worksheet) As ListObject
Set FstLo = FstItr(A.ListObjects)
End Function

