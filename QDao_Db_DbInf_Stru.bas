Attribute VB_Name = "QDao_Db_DbInf_Stru"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Db_DbInf_Stru."
Private Const Asm$ = "QDao"

Function StruInf(A As Database) As Dt
Dim T$, TT, Dy(), Des$, NRec&, Stru$
'For Each TT In TnyDb(A)
    T = TT
'    Des = Dbt_Des(A, T)
'    Stru = RmvT1(Stru(A, T))
'    NRec = NRecDT(A, T)
    PushI Dy, Array(T, NRec, Des, Stru)
'Next
StruInf = DtzFF("Tbl", "Tbl NRec Des", Dy)
End Function

Sub DmpStru(A As Database)
D Stru(A)
End Sub

Function StruFld(ParamArray Ap()) As Drs
Dim Dy(), S$, I, Av(), Ele$, LikFF$, LikFld$, J
Av = Ap
For Each I In Av
    S = I
    AsgTRst S, Ele, LikFF
    For Each J In SyzSS(LikFF)
        LikFld = J
        PushI Dy, Array(Ele, LikFld)
    Next
Next
StruFld = DrszFF("Ele FldLik", Dy)
End Function

Sub DmpStruTT(A As Database, TT$)
D StruzTT(A, TT)
End Sub
