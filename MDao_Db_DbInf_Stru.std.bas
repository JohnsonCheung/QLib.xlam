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

Function Db_Stru$(A As Database)
'Db_Stru = StruDTT(A, TnyDb(A))
End Function

Sub Db_Stru_Dmp(A As Database)
D Db_Stru(A)
End Sub

Function StruFld(ParamArray Ap()) As Drs
Dim Dry(), Av(), Ele$, LikFF, LikFld, X
Av = Ap
For Each X In Av
    AsgTRst X, Ele, LikFF
    For Each LikFld In SySsl(LikFF)
        PushI Dry, Array(Ele, LikFld)
    Next
Next
Set StruFld = Drs("Ele FldLik", Dry)
End Function

Function StruTT$(TT)
'StruTT = StruDbtt(CDb, TT)
End Function

Sub DmpStruT(TT)
D StruTT(TT)
End Sub
