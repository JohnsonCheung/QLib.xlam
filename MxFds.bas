Attribute VB_Name = "MxFds"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxFds."

Function CsvzFds$(A As dao.Fields)
CsvzFds = CsvLinzDr(AvzItr(A))
End Function

Function DrzFds(A As dao.Fields, Optional FF$) As Variant()
If FF = "" Then
    Dim F As dao.Field
    For Each F In A
        PushI DrzFds, EmptyIfNull(F.Value)
    Next
    Exit Function
End If
DrzFds = DrzFdsFny(A, Ny(FF))
End Function

Function DrzFdsFny(A As dao.Fields, Fny$()) As Variant()
Dim I, O()
For Each I In Fny
    Push O, EmptyIfNull(A(I).Value)
Next
End Function

Private Sub Z()
Z_DrzFds
MDao_Z_Fds:
End Sub

Private Sub Z_DrzFds()
Dim Rs As dao.Recordset, Dy()
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
Dim Rs As dao.Recordset, Dr(), D As Database
Set Rs = RszQ(SampDbDutyDta, "Select * from SkuB")
With Rs
    While Not .EOF
        Dr = DrzRs(Rs)
        Debug.Print JnComma(Dr)
        .MoveNext
    Wend
    .Close
End With
End Sub
