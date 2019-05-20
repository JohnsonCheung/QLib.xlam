Attribute VB_Name = "QDao_Def_Fds"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Def_Fds."
Private Const Asm$ = "QDao"

Function CsvzFds$(A As DAO.Fields)
CsvzFds = CsvzDr(AvzItr(A))
End Function

Function DrzFds(A As DAO.Fields, Optional FF$) As Variant()
If FF = "" Then
    Dim F As DAO.Field
    For Each F In A
        PushI DrzFds, EmptyIfNull(F.Value)
    Next
    Exit Function
End If
DrzFds = DrzFdsFny(A, Ny(FF))
End Function

Function DrzFdsFny(A As DAO.Fields, Fny$()) As Variant()
Dim I, O()
For Each I In Fny
    Push O, EmptyIfNull(A(I).Value)
Next
End Function

Private Sub Z_DrzFds()
Dim Rs As DAO.Recordset, Dry()
Set Rs = Db(SampFbzShpRate).OpenRecordset("Select * from YMGRnoIR")
With Rs
    While Not .EOF
        PushI Dry, DrzFds(Rs.Fields)
        .MoveNext
    Wend
    .Close
End With
BrwDry Dry
End Sub

Private Sub Z_DrzFds1()
Dim Rs As DAO.Recordset, Dr(), D As Database
Set Rs = RszQ(SampDbzDutyDta, "Select * from SkuB")
With Rs
    While Not .EOF
        Dr = DrzRs(Rs)
        Debug.Print JnComma(Dr)
        .MoveNext
    Wend
    .Close
End With
End Sub

Private Sub ZZ()
Z_DrzFds
MDao_Z_Fds:
End Sub

