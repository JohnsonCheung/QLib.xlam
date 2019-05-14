Attribute VB_Name = "QDta_Bql_WrtBql"
Option Explicit
Private Const CMod$ = "BBqlWrite."
Public Const DoczFbq$ = "Fbq is Full file name of back quote (`) separated lines. " & _
"It has first line as ShtTyscfQBLin.  " & _
"It rest of lines are records."
Sub InsRszBql(R As Dao.Recordset, Bql$)
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


Private Sub Z_WrtFbqlzDb()
Dim P$: P = TmpPth
WrtFbqlzDb P, SampDbzDutyDta
BrwPth P
Stop
End Sub

Private Sub Z_WrtFbqlzT()
Dim T$: T = TmpFt
WrtFbql T, SampDbzDutyDta, "PermitD"
BrwFt T
End Sub

Sub WrtFbqlzDb(Pth, A As Database)
WrtFbqlzTny Pth, A, Tny(A)
End Sub

Sub WrtFbqlzTny(Pth, A As Database, Tny$())
Dim T, P$
P = EnsPthSfx(Pth)
For Each T In Tny
    WrtFbql P & T & ".bql.txt", A
Next
End Sub

Sub WrtFbql(Fbql$, A As Database, Optional T0$)
Dim T$
    T = T0
    If T = "" Then T = TblNmzFbql(Fbql)
Dim F%: F = FnoO(Fbql)
Dim R As Dao.Recordset
Set R = RszT(A, T)
Dim L$: L = ShtTyBqlzT(A, T)
Print #F, L
With R
    While Not .EOF
        Print #F, BqlzRs(R)
        .MoveNext
    Wend
    .Close
End With
Close #F
End Sub
