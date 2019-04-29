Attribute VB_Name = "MDao_Tmp"
Option Explicit
Property Get TmpTd() As DAO.TableDef
Dim Fdy() As Field2
PushObj Fdy, FdzTxt("F1")
Set TmpTd = TdzFdy("Tmp", Fdy)
End Property

Property Get TmpDbPth$()
Dim O$: O = TmpHom & "Db\"
EnsPth O
TmpDbPth = O
End Property

Function TmpDb(Optional Fdr$, Optional Fnn$) As Database
Dim Fb$: Fb = TmpFb
CrtFb Fb
Set TmpDb = Db(Fb$)
End Function

Function TmpFb$()
TmpFb = TmpDbPth & TmpNm & ".accdb"
End Function


