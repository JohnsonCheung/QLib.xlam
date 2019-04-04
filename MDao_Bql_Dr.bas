Attribute VB_Name = "MDao_Bql_Dr"
Option Explicit
Public Const DocOfBql$ = "is Back quote (`) separated line.  If the field is blank, don't set Rs's value"
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
If L = "401`HD0V4FOF00C9ZT" Then Stop
BqlzRs = L

End Function


