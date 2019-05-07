Attribute VB_Name = "QDao_Tmp"
Option Explicit
Private Const CMod$ = "MDao_Tmp."
Private Const Asm$ = "QDao"
Property Get TmpTd() As DAO.TableDef
Dim Fdy() As DAO.Field2
PushObj Fdy, FdzTxt("F1")
Set TmpTd = TdzFdy("Tmp", Fdy)
End Property

Property Get TmpPthzDb$()
Dim O$: O = TmpHom & "Db\"
EnsPth O
TmpPthzDb = O
End Property

Function TmpDb(Optional Fdr$, Optional Fnn$) As Database
Dim Fb$: Fb = TmpFb
CrtFb Fb
Set TmpDb = Db(Fb$)
End Function

Function TmpFb$()
TmpFb = TmpPthzDb & TmpNm & ".accdb"
End Function


