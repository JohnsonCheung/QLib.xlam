Attribute VB_Name = "QDao_Lnk"
Option Explicit
Private Const CMod$ = "MDao_Lnk."
Private Const Asm$ = "QDao"

Sub LnkTbl(A As Database, T, S$, Cn$)
On Error GoTo X
DrpT A, T
A.TableDefs.Append TdzTSCn(T, S, Cn)
Exit Sub
X:
    Dim Er$: Er = Err.Description
    Thw CSub, "Error in linking table", "Er Db T SrcTbl Cn", Er, Dbn(A), T, S, Cn
End Sub

Sub LnkFxw(A As Database, T, Fx, Optional Wsn = "Sheet1")
LnkTbl A, T, Wsn & "$", CnStrzFxDAO(Fx)
End Sub

Sub LnkFb(A As Database, T, Fb, Optional Fbt$)
Dim Cn$: Cn = CnStrzFbzAsDao(Fb)
ThwIf_Er ErzLnkTblzTSrcCn(A, T, IIf(Fbt = "", T, Fbt), Cn), CSub
End Sub

