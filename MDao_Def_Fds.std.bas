Attribute VB_Name = "MDao_Def_Fds"
Option Explicit

Function CsvzFds$(A As DAO.Fields)
CsvzFds = CsvzDr(VyzItr(A))
End Function
Function NzEmpty(A)
If IsNull(A) Then Exit Function
Asg A, NzEmpty
End Function

Function DrzFds(A As DAO.Fields, Optional FF = "") As Variant()
DrzFds = VyzFds(A, FF)
End Function

Function FnyzFds(A As Fields) As String()
FnyzFds = Itn(A)
End Function

Function VyzFds(A As DAO.Fields, Optional FF = "") As Variant()
Dim J%, F As DAO.Field, N
If FF = "" Then
    For Each F In A
        PushI VyzFds, F.Value
    Next
Else
    For Each N In FnyzFF(FF)
        Push VyzFds, A(F).Value
    Next
End If
End Function

Private Sub Z_DrzFds()
Dim Rs As DAO.Recordset, Dry()
Set Rs = Db(SampFbzShpRate).OpenRecordset("Select * from YMGRnoIR")
With Rs
    While Not .EOF
        Push Dry, DrzFds(Rs.Fields)
        .MoveNext
    Wend
    .Close
End With
Brw FmtDry(Dry)
End Sub

Private Sub Z_VyzFds()
Dim Rs As DAO.Recordset, Vy()
'Set Rs = CDb.OpenRecordset("Select * from SkuB")
With Rs
    While Not .EOF
        Vy = VyzFds(Rs)
        Debug.Print JnComma(Vy)
        .MoveNext
    Wend
    .Close
End With
End Sub



Private Sub Z()
Z_DrzFds
Z_VyzFds
MDao_Z_Fds:
End Sub
