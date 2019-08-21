Attribute VB_Name = "QVb_Str_Box"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Str_Box."
Private Const Asm$ = "QVb"

Function BoxLines(Lines, Optional C$ = "*") As String()
BoxLines = BoxLy(SplitCrLf(Lines))
End Function
Function BoxLy(Ly$(), Optional C$ = "*") As String()
If Si(Ly) = 0 Then Exit Function
Dim W%, L$, I
W = AyWdt(Ly)
L = Qte(Dup("-", W), "|-*-|")
PushI BoxLy, L
For Each I In Ly
    PushI BoxLy, "| " & AlignL(I, W) & " |"
Next
PushI BoxLy, L
End Function
Function BoxS(S, Optional C$ = "*") As String()
Dim H$: H = Dup(C, Len(S) + 6)
PushI BoxS, H
PushI BoxS, C & C & " " & S & " " & C & C
PushI BoxS, H
End Function
Function Box(V, Optional C$ = "*") As String()
If IsStr(V) Then
    If V = "" Then
        Exit Function
    End If
End If
Select Case True
Case IsLines(V): Box = BoxLines(V, C)
Case IsStr(V):   Box = BoxS(V, C)
Case IsSy(V):    Box = BoxLy(CvSy(Sy), C)
Case IsArray(V): Box = BoxAy(V)
Case Else:       Box = BoxS(V, C)
End Select
End Function

Function BoxFny(Fny$()) As String()
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
BoxFny = Sy(H, L, H)
End Function

Function BoxAy(Ay) As String()
If Si(Ay) = 0 Then Exit Function
Dim W%: W = AyWdt(Ay)
Dim H$: H = "|" & Dup("-", W + 2) & "|"
Push BoxAy, H
Dim I
For Each I In Ay
    Push BoxAy, "| " & AlignL(I, W) + " |"
Next
Push BoxAy, H
End Function
