Attribute VB_Name = "MDao_Tmp"
Option Explicit
Property Get TmpTd() As Dao.TableDef
Dim X() As Dao.Field2
PushObj X, FdzTxt("F1")
Set TmpTd = TdzFdAy("Tmp", X)
End Property

Property Get TmpDbPth$()
TmpDbPth = PthEns(TmpHom & "Db\")
End Property

Function TmpDb(Optional Fdr$, Optional Fnn$) As Database
Dim Fb$: Fb = TmpFb
CrtFb Fb
Set TmpDb = Db(Fb)
End Function

Function TmpFb$()
TmpFb = TmpDbPth & TmpNm & ".accdb"
End Function


