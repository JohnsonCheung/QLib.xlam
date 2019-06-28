Attribute VB_Name = "QDao_Def_Fds"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Def_Fds."
Private Const Asm$ = "QDao"

Function CsvzFds$(A As Dao.Fields)
CsvzFds = CsvzDr(AvzItr(A))
End Function

Function DrzFds(A As Dao.Fields, Optional FF$) As Variant()
If FF = "" Then
    Dim F As Dao.Field
    For Each F In A
        PushI DrzFds, EmptyIfNull(F.Value)
    Next
    Exit Function
End If
DrzFds = DrzFdsFny(A, Ny(FF))
End Function

Function DrzFdsFny(A As Dao.Fields, Fny$()) As Variant()
Dim I, O()
For Each I In Fny
    Push O, EmptyIfNull(A(I).Value)
Next
End Function

Private Sub Z_DrzFds()
Dim Rs As Dao.Recordset, Dy()
Set Rs = Db(SampFbzShpRate).OpenRecordset("Select * from YMGRnoIR")
With Rs
    While Not .EOF
        PushI Dy, DrzFds(Rs.Fields)
        .MoveNext
    Wend
    .Close
End With
BrwDy Dy
End Sub

Private Sub Z_DrzFds1()
Dim Rs As Dao.Recordset, Dr(), D As Database
Set Rs = RszQ(SampDboDutyDta, "Select * from SkuB")
With Rs
    While Not .EOF
        Dr = DrzRs(Rs)
        Debug.Print JnComma(Dr)
        .MoveNext
    Wend
    .Close
End With
End Sub

Private Sub Z()
Z_DrzFds
MDao_Z_Fds:
End Sub

