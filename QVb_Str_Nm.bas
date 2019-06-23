Attribute VB_Name = "QVb_Str_Nm"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Str_Nm."
Private Const Asm$ = "QVb"
Public Const C_Dbl$ = """"
Public Const C_Sng$ = "'"

Function IsNm(S) As Boolean
If S = "" Then Exit Function
If Not IsLetter(FstChr(S)) Then Exit Function
Dim L&: L = Len(S)
If L > 64 Then Exit Function
Dim J%
For J = 2 To L
   If Not IsChrzNm(Mid(S, J, 1)) Then Exit Function
Next
IsNm = True
End Function

Function IsChrzNm(C$) As Boolean
IsChrzNm = True
If IsLetter(C) Then Exit Function
If C = "_" Then Exit Function
If IsDigit(C) Then Exit Function
IsChrzNm = False
End Function

Function IsChrDotNm(A$) As Boolean
If IsChrzNm(A) Then IsChrDotNm = True: Exit Function
IsChrDotNm = A = "."
End Function

Function WhNmzS(WhStr$) As WhNm
Dim P As Dictionary: Set P = Lpm(WhStr, "-Sw Prv Pub Frd Sub Fun Prp Get Set Let WiRet WoRet")
'WhNmzS = WhNmzP(P,WhNm(NmPfx)
End Function

Function QtezChr$(S, Chr$)
QtezChr = Chr & S & Chr
End Function

Function QtezSpc$(S)
QtezSpc = QtezChr(S, " ")
End Function

Function QtezDblQ$(S)
QtezDblQ = QtezChr(S, vbQtezDblQ)
End Function

Function SngQte$(S)
SngQte = QtezChr(S, vbSngQte)
End Function

Function HitRe(S, Re As RegExp) As Boolean
If S = "" Then Exit Function
If IsNothing(Re) Then Exit Function
HitRe = Re.Test(S)
End Function

Function NmSfx$(S)
Dim J%, O$, C$
For J = Len(S) To 1 Step -1
    C = Mid(S, J, 1)
    If Not IsAscUCas(Asc(C)) Then
        If C <> "_" Then
            NmSfx = O: Exit Function
        End If
    End If
    O = C & O
Next
End Function

Function NxtSeqNm$(Nm$, Optional NDig& = 3) _
'Nm-Nm can be XXX or XXX_nn
'   If XXX, return XXX_001   '<-- # of zero depends on NDig
'   If XXX_nn, return XXX_mm '<-- mm is nn+1, # of digit of nn and mm depends on NDig
If NDig = 0 Then Stop
Dim R$
    R = Right(Nm, NDig + 1)

If Left(R, 1) <> "_" Then GoTo Case1
If Not IsNumeric(Mid(R, 2)) Then GoTo Case1

Dim L$: L = Left(Nm, Len(Nm) - NDig)
Dim Nxt&: Nxt = Val(Mid(R, 2)) + 1
NxtSeqNm = Left(Nm, Len(Nm) - NDig) + Pad0(Nxt, NDig)
Exit Function

Case1:
    NxtSeqNm = Nm & "_" & Dup(NDig - 1, "0") & "1"
End Function



