Attribute VB_Name = "MVb_Str_Nm"
Option Explicit
Public Const C_Dbl$ = """"
Public Const C_Sng$ = "'"
Function NmSeqNo%(A)
Dim B$: B = TakAftRev(A, "_")
If B = "" Then Exit Function
If Not IsNumeric(B) Then Exit Function
NmSeqNo = B
End Function

Sub DDNmAsgBrk(A, O1$, O2$, O3$)
Dim Ay$(): Ay = Split(A, ".")
Select Case Sz(Ay)
Case 1: O1 = "":    O2 = "":    O3 = Ay(0)
Case 2: O1 = "":    O2 = Ay(0): O3 = Ay(1)
Case 3: O1 = Ay(0): O2 = Ay(1): O3 = Ay(2)
Case Else: Stop
End Select
End Sub


Function DDNmThird$(A)
Dim Ay$(): Ay = Split(A, "."): If Sz(Ay) <> 3 Then Stop
DDNmThird = Ay(2)
End Function

Function IsNm(A) As Boolean
If Not IsLetter(FstChr(A)) Then Exit Function
Dim L%: L = Len(A)
If L > 64 Then Exit Function
Dim J%
For J = 2 To L
   If Not IsNmChr(Mid(A, J, 1)) Then Exit Function
Next
IsNm = True
End Function

Function IsNmChr(A$) As Boolean
IsNmChr = True
If IsLetter(A) Then Exit Function
If A = "_" Then Exit Function
If IsDigit(A) Then Exit Function
IsNmChr = False
End Function
Function WhNmzStr(WhStr$, Optional NmPfx$) As WhNm
Set WhNmzStr = LinPm(WhStr).WhNm(NmPfx)
End Function

Function ChrQuote$(S, Chr$)
ChrQuote = Chr & S & Chr
End Function
Function SpcQuote$(S)
SpcQuote = ChrQuote(S, " ")
End Function
Function DblQuote$(S)
DblQuote = ChrQuote(S, vbDblQuote)
End Function
Function SngQuote$(S)
SngQuote = ChrQuote(S, vbSngQuote)
End Function

Function HitRe(Str, Re As RegExp) As Boolean
If Str = "" Then Exit Function
If IsNothing(Re) Then Exit Function
HitRe = Re.Test(Str)
End Function

Function NmSfx$(A)
Dim J%, O$, C$
For J = Len(A) To 1 Step -1
    C = Mid(A, J, 1)
    If Not IsAscUCase(Asc(C)) Then
        If C <> "_" Then
            NmSfx = O: Exit Function
        End If
    End If
    O = C & O
Next
End Function

Function NxtSeqNm$(Nm, Optional NDig% = 3) _
'Nm-Nm can be XXX or XXX_nn
'   If XXX, return XXX_001   '<-- # of zero depends on NDig
'   If XXX_nn, return XXX_mm '<-- mm is nn+1, # of digit of nn and mm depends on NDig
If NDig = 0 Then Stop
Dim R$
    R = Right(Nm, NDig + 1)

If Left(R, 1) <> "_" Then GoTo Case1
If Not IsNumeric(Mid(R, 2)) Then GoTo Case1

Dim L$: L = Left(Nm, Len(Nm) - NDig)
Dim Nxt%: Nxt = Val(Mid(R, 2)) + 1
NxtSeqNm = Left(Nm, Len(Nm) - NDig) + Pad0(Nxt, NDig)
Exit Function

Case1:
    NxtSeqNm = Nm & "_" & Dup(NDig - 1, "0") & "1"
End Function



