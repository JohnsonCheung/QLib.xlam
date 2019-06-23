Attribute VB_Name = "QDta_Bql_WrtBql"
Option Compare Text
Option Explicit
Private Const CMod$ = "BBqlWrite."
':Fbka: # :Ft:       ! #Fil-BacK-Apostrophe#  It is a fil of ext *.bka.txt.  There 1-Sgn-Lin, 0-to-N-Rmk-Lines and 0-To-N-Tbl-Lines.
'Sgn-Lin             ! is  **BackApostropheSeparatedFile**<Dsn>**, where <Dsn> is a dta-set-nm.
'Rmk-Lines           ! are lines between Sgn-Lin and (fst-TblNm-Lin or eof)
'1-Tbl-Lines         ! is  1-TblNm-Lin, 0-to-N-TblRmk-Lines, 1-Fld-Lin and 0-to-N-Dta-Lines.
'TblNm-Lin           ! is
'                    ! Rmk-Lines are lines before fst *-Lin.  Rmk are for all tbl in the :Fbka:  Each individual tbl does not have it own rmk
'Lines before fst *-lin are Rmk.  Each gp of one-*-Lin & N-`-Lin is one tbl.
'                    ! *-lin is a lin wi fst chr is *, :Starl: #Star-Line#.  `-lin is a lin wi fst chr is `, :Bkal:, #BacK-Apostrophe-Lin#.
'                    ! The *-Lin is *<Tn>
'                    ! The fst `-Lin is :Scff:
'                    ! The rst `-Lin is :dta:
':Scff: # :SS:Sc
':Tn:   # :s:        ! #Table-Name#.
':Scff: # :SS:Scfld: ! #ShtTyChr-Colon-FF#.  It is spc sep of :Scfld:.  It desc ty and fldn of the tbl.
'It has first line as ShtTyscfQBLin.  " & _
"It rest of lines are records."

Sub InsRszBql(R As DAO.Recordset, Bql$)
R.AddNew
Dim Ay$(): Ay = Split(Bql, "`")
Dim F As DAO.Field, J%
For Each F In R.Fields
    If Ay(J) <> "" Then
        F.Value = Ay(J)
    End If
    J = J + 1
Next
R.Update
End Sub
Function BqlzRs$(A As DAO.Recordset)
Dim O$(), F As DAO.Field
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
Dim R As DAO.Recordset
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
