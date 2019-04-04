Attribute VB_Name = "MTp_SqyRslt_31_SwBrkAy"
Option Explicit

Private Function SwBrk(A As Lnx) As SwBrk
Dim L$, Ix%, OEr$()
L = A.Lin
Ix = A.Ix
If IsDDRmkLin(L) Then Thw CSub, "[SwLin], [Ix] is a remark line.  It should be removed before calling Evl", A.Lin, A.Ix
Set SwBrk = New SwBrk
With SwBrk
    .Nm = ShfT1(L)
    .OpStr = UCase(ShfT1(L))
    .TermAy = SySsl(L)
    .Ix = Ix
End With
End Function

Function SwBrkAy(A() As Lnx) As SwBrk()
Dim I
For Each I In Itr(A)
    PushObj SwBrkAy, SwBrk(CvLnx(I))
Next
End Function

