Attribute VB_Name = "QDao_Db_DbInf_Stru"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Db_DbInf_Stru."
Private Const Asm$ = "QDao"

Function StruInf(A As Database) As Dt
Dim T$, TT, Dry(), Des$, NRec&, Stru$
'For Each TT In TnyDb(A)
    T = TT
'    Des = Dbt_Des(A, T)
'    Stru = RmvT1(Stru(A, T))
'    NRec = NRecDT(A, T)
    PushI Dry, Array(T, NRec, Des, Stru)
'Next
StruInf = DtzFF("Tbl", "Tbl NRec Des", Dry)
End Function

Sub DmpStru(A As Database)
D Stru(A)
End Sub

Function StruFld(ParamArray Ap()) As Drs
Dim Dry(), S$, I, Av(), ele$, LikFF$, LikFld$, J
Av = Ap
For Each I In Av
    S = I
    AsgTRst S, ele, LikFF
    For Each J In SyzSS(LikFF)
        LikFld = J
        PushI Dry, Array(ele, LikFld)
    Next
Next
StruFld = DrszFF("Ele FldLik", Dry)
End Function

Sub DmpStruTT(A As Database, TT$)
D StruzTT(A, TT)
End Sub
