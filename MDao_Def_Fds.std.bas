Attribute VB_Name = "MDao_Def_Fds"
Option Explicit

Function CsvzFds$(A As Dao.Fields)
CsvzFds = CsvzDr(VyzItr(A))
End Function
Function NzEmpty(A)
If IsNull(A) Then Exit Function
Asg A, NzEmpty
End Function

Function DrzFds(A As Dao.Fields, Optional FF = "") As Variant()
DrzFds = VyzFds(A, FF)
End Function

Function FnyzFds(A As Fields) As String()
FnyzFds = Itn(A)
End Function

Function VyzFds(A As Dao.Fields, Optional FF = "") As Variant()
Dim F As Dao.Field, N, O()
If FF = "" Then
    For Each F In A
        PushI O, F.Value
    Next
Else
    For Each N In FnyzFF(FF)
        Push O, A(F).Value
    Next
End If
Dim J%
Dim I
For Each I In O
    If IsNull(I) Then
        O(J) = ""
    End If
    J = J + 1
Next
VyzFds = O
End Function

Private Sub Z_DrzFds()
Dim Rs As Dao.Recordset, Dry()
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
Dim Rs As Dao.Recordset, Vy(), D As Database
'Set Rs = D.OpenRecordset("Select * from SkuB")
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
