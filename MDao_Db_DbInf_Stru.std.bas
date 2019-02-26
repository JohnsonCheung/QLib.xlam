Attribute VB_Name = "MDao_Db_DbInf_Stru"
Option Explicit

Function DbInfDtStru(A As Database) As Dt
Dim T$, TT, Dry(), Des$, NRec&, Stru$
'For Each TT In TnyDb(A)
    T = TT
'    Des = Dbt_Des(A, T)
'    Stru = RmvT1(Stru(A, T))
'    NRec = NRecDT(A, T)
    PushI Dry, Array(T, NRec, Des, Stru)
'Next
Set DbInfDtStru = Dt("Tbl", "Tbl NRec Des", Dry)
End Function

Sub DmpStru(A As Database)
D Stru(A)
End Sub

Function StruFld(ParamArray Ap()) As DRs
Dim Dry(), Av(), Ele$, LikFF, LikFld, X
Av = Ap
For Each X In Av
    AsgTRst X, Ele, LikFF
    For Each LikFld In SySsl(LikFF)
        PushI Dry, Array(Ele, LikFld)
    Next
Next
Set StruFld = DRs("Ele FldLik", Dry)
End Function

Sub DmpStruTT(A As Database, TT)
D StruzTT(A, TT)
End Sub
