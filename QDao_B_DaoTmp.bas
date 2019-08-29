Attribute VB_Name = "QDao_B_DaoTmp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Tmp."
Private Const Asm$ = "QDao"

Sub BrwLasTmpDb()
BrwDb LasTmpDb
End Sub

Function LasTmpDb() As Database
Set LasTmpDb = Db(LasTmpFb)
End Function

Function LasTmpFb$()
Dim Fn$: Fn = MaxEle(FnAy(TmpDbPth, "*.accdb"))
If Fn = "" Then Thw CSub, "No *.accdb TmpDbPth", "TmpDbPth", TmpDbPth
LasTmpFb = TmpDbPth & Fn
End Function

Function TmpDb() As Database
Dim Fb$: Fb = TmpFb
CrtFb Fb
Set TmpDb = Db(Fb)
End Function

Property Get TmpDbPth$()
TmpDbPth = EnsPth(TmpHom & "TmpDb\")
End Property

Function TmpFb$()
TmpFb = TmpDbPth & TmpNm & ".accdb"
End Function
