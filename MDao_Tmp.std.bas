Attribute VB_Name = "MDao_Tmp"
Option Explicit
Property Get TmpTd() As DAO.TableDef
Dim O() As DAO.Field2
PushObj O, NewFd("F1")
'Set TmpTd = NewTd("Tmp", O)
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


