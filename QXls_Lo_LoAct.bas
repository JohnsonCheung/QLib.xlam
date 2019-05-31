Attribute VB_Name = "QXls_Lo_LoAct"
Option Compare Text
Option Explicit
Private Const Asm$ = "QXls"
Private Const CMod$ = "MXls_Lo_Get_Prp."

Sub BrwLo(A As ListObject)
BrwDrs DrszLo(A)
End Sub

Sub MinxLo(A As ListObject)
If FstTwoChr(A.Name) <> "T_" Then Exit Sub
Dim R1 As Range
Set R1 = A.DataBodyRange
If R1.Rows.Count >= 2 Then
    RgRR(R1, 2, R1.Rows.Count).EntireRow.Delete
End If
End Sub

Private Sub MinxLozWs(A As Worksheet)
If A.CodeName = "WsIdx" Then Exit Sub
If FstTwoChr(A.CodeName) <> "Ws" Then Exit Sub
Dim L As ListObject
For Each L In A.ListObjects
    MinxLo L
Next
End Sub

Function MinxLozWszWb(A As Workbook) As Workbook
Dim Ws As Worksheet
For Each Ws In A.Sheets
    MinxLozWs Ws
Next
Set MinxLozWszWb = A
End Function

Sub MinxLozWszFx(Fx)
ClsWbNoSav SavWb(MinxLozWszWb(WbzFx(Fx)))
End Sub


Sub KeepFstCol(A As ListObject)
Dim J%
For J = A.ListColumns.Count To 2 Step -1
    A.ListColumns(J).Delete
Next
End Sub

Sub KeepFstRow(A As ListObject)
Dim J%
For J = A.ListRows.Count To 2 Step -1
    A.ListRows(J).Delete
Next
End Sub


Function LoPc(A As ListObject) As PivotCache
Dim O As PivotCache
Set O = WbzLo(A).PivotCaches.Create(xlDatabase, A.Name, 6)
O.MissingItemsLimit = xlMissingItemsNone
Set LoPc = O
End Function

Function LoQt(A As ListObject) As QueryTable
On Error Resume Next
Set LoQt = A.QueryTable
End Function

Function R1Lo&(A As ListObject, Optional InclHdr As Boolean)
If IsLozNoDta(A) Then
   R1Lo = A.ListColumns(1).Range.Row + 1
   Exit Function
End If
R1Lo = A.DataBodyRange.Row - IIf(InclHdr, 1, 0)
End Function

Function R2Lo&(A As ListObject, Optional InclTot As Boolean)
If IsLozNoDta(A) Then
   R2Lo = R1Lo(A)
   Exit Function
End If
R2Lo = A.DataBodyRange.Row + IIf(InclTot, 1, 0)
End Function

Function SqzLo(A As ListObject) As Variant()
SqzLo = A.DataBodyRange.Value
End Function

Function WszLo(A As ListObject) As Worksheet
Set WszLo = A.Parent
End Function

Function WbzLo(A As ListObject) As Workbook
Set WbzLo = WbzWs(WszLo(A))
End Function

Function WsCnozLc&(A As ListObject, Col)
WsCnozLc = A.ListColumns(Col).Range.Column
End Function

Function LoNmzT$(T)
LoNmzT = "T_" & RmvFstNonLetter(T)
End Function

Private Sub ZZ_KeepFstCol()
Dim Lo As ListObject
KeepFstCol ShwLo(SampLo)
End Sub

Private Sub Z_AutoFitLo()
Dim Ws As Worksheet: Set Ws = NewWs
Dim Sq(1 To 2, 1 To 2)
Sq(1, 1) = "A"
Sq(1, 2) = "B"
Sq(2, 1) = "123123"
Sq(2, 2) = String(1234, "A")
Ws.Range("A1:B2").Value = Sq
AutoFitLo LozWsDta(Ws)
ClsWsNoSav Ws
End Sub

Private Sub Z_BrwLo()
BrwLo SampLo
Stop
End Sub

Private Sub Z_PtzLo()
Dim At As Range, Lo As ListObject
Set Lo = SampLo
ShwPt PtzLo(Lo, At, "A B", "C D", "F", "E")
Stop
End Sub

Private Sub ZZ()
Dim A As Variant
Dim B As ListObject
Dim C As Boolean
Dim D As Worksheet
Dim E$
Dim F&
Dim G$()
Dim H As Range
CvLo A
LoAllCol B
LoAllEntCol B
AutoFitLo B
BdrLoAround B
RgzLoCC B, A, A, C, C
RgzLc B, A, C, C
RgzLc B, A
SqzLo B
ShwLo B
WbzLo B
WszLo B
End Sub

