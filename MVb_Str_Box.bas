Attribute VB_Name = "MVb_Str_Box"
Option Explicit

Function BoxLyLines(Lines$) As String()
BoxLyLines = BoxLyAy(SplitCrLf(Lines))
End Function

Function BoxLyAy(Ay) As String()
If Si(Ay) = 0 Then Exit Function
Dim W%: W = AyWdt(Ay)
Dim H$: H = "|" & Dup("-", W + 2) & "|"
Push BoxLyAy, H
Dim I
For Each I In Ay
    Push BoxLyAy, "| " & AlignL(I, W) + " |"
Next
Push BoxLyAy, H
End Function
