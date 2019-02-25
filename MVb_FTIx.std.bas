Attribute VB_Name = "MVb_FTIx"
Option Explicit

Function FTIx_HasU(A As FTIx, U&) As Boolean
If U < 0 Then Stop
If A.IsEmp Then Exit Function
If A.FmIx > U Then Exit Function
If A.ToIx < U Then Exit Function
FTIx_HasU = True
End Function

Sub AssBet(Fun$, V, FmV, ToV)
If FmV > V Then Thw Fun, "FmV > V", "V FmV ToV", V, FmV, ToV
If ToV < V Then Thw Fun, "ToV < V", "V FmV ToV", V, FmV, ToV
End Sub

Function FTIxLinCnt%(A As FTIx)
Dim O%
O = A.ToIx - A.FmIx + 1
If O < 0 Then Stop
FTIxLinCnt = O
End Function

Function FTIxFTIx(A As FTIx) As FTIx
Dim Lno&, Cnt&
   Cnt = A.ToIx - A.FmIx + 1
   If Cnt < 0 Then Cnt = 0
   Lno = A.FmIx + 1
Set FTIxFTIx = FTIx(Lno, Cnt)
End Function

Function FTIx(FmIx, ToIx) As FTIx
Dim O As New FTIx
Set FTIx = O.Init(FmIx, ToIx)
End Function
Function CvFTIx(A) As FTIx
Set CvFTIx = A
End Function


