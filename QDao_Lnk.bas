Attribute VB_Name = "QDao_Lnk"
Option Explicit
Private Const CMod$ = "MDao_Lnk."
Private Const Asm$ = "QDao"

Sub LnkTbl(A As Database, T$, S$, Cn$)
On Error GoTo X
DrpT A, T
A.TableDefs.Append TdzTSCn(T, S, Cn)
Exit Sub
X:
    Dim Er$: Er = Err.Description
    Thw CSub, "Error in linking table", "Er Db T SrcTbl Cn", Er, DbNm(A), T, S, Cn
End Sub

Function LnkFxw(A As Database, T$, Fx$, Optional Wsn = "Sheet1") As String()
LnkFxw = LnkTbl(A, T, Wsn & "$", CnStrzFxDAO(Fx$))
End Function

Function LnkFbtt(A As Database, TTCrt$, Fb$, Optional Fbtt$) As String()
Dim TnyCrt$(), TnyzFb$(), J%, T
TnyCrt = TermSy(TTCrt)
TnyzFb = IIf(Fbtt = "", TnyCrt, TermSy(Fbtt))
If Si(TnyzFb) <> Si(TnyCrt) Then
    Thw CSub, "[TTCrt] and [FbttSz] are diff", "TTCrtSz FbttSz TnyCrt TnyzFb GivenFbtt", Si(TnyCrt), Si(TnyzFb), TnyCrt, TnyzFb, Fbtt
End If
Dim Cn$: Cn = CnStrzFbzAsDao(Fb$)
For J = 0 To UB(TnyCrt)
    PushIAy LnkFbtt, LnkTbl(A, TnyCrt(J), TnyzFb(J), Cn)
Next
End Function
Sub LnkFb(A As Database, T$, Fb$, Optional Fbt$)
Dim Cn$: Cn = CnStrzFbzAsDao(Fb$)
ThwIfEr ErzLnkTblzTSrcCn(A, T, IIf(Fbt = "", T, Fbt), Cn), CSub
End Sub

