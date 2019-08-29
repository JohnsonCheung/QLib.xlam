Attribute VB_Name = "QDao_Db_Stru"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Db_DbInf_Stru."
Private Const Asm$ = "QDao"

Sub DmpStru(D As Database)
Dmp Stru(D)
End Sub

Sub DmpStruTT(D As Database, TT$)
Dmp StruzTT(D, TT)
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

Function StruInf(D As Database) As DT
Dim T$, TT, Dy(), Des$, NRec&, Stru$
'For Each TT In TnyDb(D)
    T = TT
'    Des = Dbt_Des(D, T)
'    Stru = RmvT1(Stru(D, T))
'    NRec = NRecDT(D, T)
    PushI Dy, Array(T, NRec, Des, Stru)
'Next
StruInf = DtzFF("Tbl", "Tbl NRec Des", Dy)
End Function
