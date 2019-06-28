Attribute VB_Name = "QVb_Str_Box"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Str_Box."
Private Const Asm$ = "QVb"

Function BoxzLines(Lines$) As String()
BoxzLines = BoxzLy(SplitCrLf(Lines))
End Function
Function BoxzLy(Ly$()) As String()
If Si(Ly) = 0 Then Exit Function
Dim W%, L$, I
W = WdtzAy(Ly)
L = Qte(Dup("-", W), "|-*-|")
PushI BoxzLy, L
For Each I In Ly
    PushI BoxzLy, "| " & AlignL(I, W) & " |"
Next
PushI BoxzLy, L
End Function
Function BoxzS(S$) As String()
Dim H$: H = Dup("*", Len(S) + 6)
PushI BoxzS, H
PushI BoxzS, "** " & S & " **"
PushI BoxzS, H
End Function
Function Box(V) As String()
If IsStr(V) Then
    If V = "" Then
        Exit Function
    End If
End If
Select Case True
Case IsLines(V): Box = BoxzLines(CStr(V))
Case IsStr(V): Box = BoxzS(CStr(V))
Case IsSy(V): Box = BoxzLy(CvSy(Sy))
Case IsArray(V): Box = BoxzAy(V)
Case Else: Box = BoxzS(CStr(V))
End Select
End Function

Function BoxzFny(Fny$()) As String()
If Si(Fny) = 0 Then Exit Function
Const S$ = " | ", Q$ = "| * |"
Const LS$ = "-|-", LQ$ = "|-*-|"
Dim L$, H$, Ay$(), J%
    ReDim Ay(UB(Fny))
    For J = 0 To UB(Fny)
        Ay(J) = Dup("-", Len(Fny(J)))
    Next
L = Qte(Jn(Fny, S), Q)
H = Qte(Jn(Ay, LS), LQ)
BoxzFny = Sy(H, L, H)
End Function

Function BoxzAy(Ay) As String()
If Si(Ay) = 0 Then Exit Function
Dim W%: W = WdtzAy(Ay)
Dim H$: H = "|" & Dup("-", W + 2) & "|"
Push BoxzAy, H
Dim I
For Each I In Ay
    Push BoxzAy, "| " & AlignL(I, W) + " |"
Next
Push BoxzAy, H
End Function
