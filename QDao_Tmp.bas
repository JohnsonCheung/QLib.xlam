Attribute VB_Name = "QDao_Tmp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Tmp."
Private Const Asm$ = "QDao"

Property Get TmpTd() As DAO.TableDef
Dim Fdy() As DAO.Field2
PushObj Fdy, FdzTxt("F1")
Set TmpTd = TdzTF("Tmp", Fdy)
End Property
Function TdzTF(T, Fdy() As DAO.Field2) As DAO.TableDef
Dim O As New TableDef
O.Name = T
AddFdy O, Fdy
Set TdzTF = O
End Function
Property Get TmpDbPth$()
TmpDbPth = EnsPth(TmpHom & "TmpDb\")
End Property

Function TmpDb() As Database
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
Dim Fn$: Fn = MaxEle(FnAy(TmpDbPth, "*.accdb"))
If Fn = "" Then Thw CSub, "No *.accdb TmpDbPth", "TmpDbPth", TmpDbPth
LasTmpFb = TmpDbPth & Fn
End Function

Function TmpFb$()
TmpFb = TmpDbPth & TmpNm & ".accdb"
End Function


