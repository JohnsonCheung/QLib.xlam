Attribute VB_Name = "QDao_Tmp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Tmp."
Private Const Asm$ = "QDao"

Property Get TmpTd() As DAO.TableDef
Dim Fdy() As Field2
PushObj Fdy, FdzTxt("F1")
Set TmpTd = TdzTF("Tmp", Fdy)
End Property

Property Get TmpPthzDb$()
TmpPthzDb = EnsPth(TmpHom & "TmpDb\")
End Property

Function TmpDb(Optional Fdr$, Optional Fnn$) As Database
Dim Fb$: Fb = TmpFb
CrtFb Fb
Set TmpDb = Db(Fb)
End Function
Function LasTmpDb() As Database
Set LasTmpDb = Db(LasTmpFb)
End Function
Sub BrwLasTmpDb()
BrwDb LasTmpDb
End Sub
Function LasTmpFb$()
Dim Fn$: Fn = MaxEle(FnAy(TmpPthzDb, "*.accdb"))
If Fn = "" Then Thw CSub, "No *.accdb TmpDbPth", "TmpDbPth", TmpPthzDb
LasTmpFb = TmpPthzDb & Fn
End Function

Function TmpFb$()
TmpFb = TmpPthzDb & TmpNm & ".accdb"
End Function


