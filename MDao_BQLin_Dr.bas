Attribute VB_Name = "MDao_BQLin_Dr"
Option Explicit
Sub InsRszBql(R As Dao.Recordset, Bql)
R.AddNew
Dim Ay$(): Ay = Split(Bql, "`")
Dim F As Dao.Field, J%
For Each F In R.Fields
    If Ay(J) <> "" Then
        F.Value = Ay(J)
    End If
    J = J + 1
Next
R.Update
End Sub
Function BqlzRs$(A As Dao.Recordset)
Dim O$(), F As Dao.Field
For Each F In A.Fields
    If IsNull(F.Value) Then
        PushI O, ""
    Else
        PushI O, Replace(Replace(F.Value, vbCr, ""), vbLf, " ")
    End If
Next
Dim L$: L = Jn(O, "`")
If L = "401`HD04VFOF00C9ZT" Then Stop
BqlzRs = L

End Function


Private Function DoczBql() As String()
Erase XX
X "Bql is Back quote (`) separated line"
X ". If the field is blank, don't set Rs's value"
End Function
