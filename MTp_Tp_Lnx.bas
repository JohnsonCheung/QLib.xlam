Attribute VB_Name = "MTp_Tp_Lnx"
Option Explicit

Function CvLnx(A) As Lnx
Set CvLnx = A
End Function
Function EmpLnx() As Lnx
Static X As New Lnx, Y As Boolean
If Not Y Then Y = True: X.Ix = -1
Set EmpLnx = X
End Function
Function Lnx(Ix, Lin) As Lnx
Set Lnx = New Lnx
With Lnx
    .Lin = Lin
    .Ix = Ix
End With
End Function

Sub LnxAsg(A As Lnx, OLin$, OIx%)
With A
    OLin = .Lin
    OIx = .Ix
End With
End Sub

Sub BrwLnxAy(A() As Lnx)
BrwAy LyzLnxAyzWithLno(A)
End Sub

Function DupT2AyzLnxAy(A() As Lnx) As String()
DupT2AyzLnxAy = AywDup(T2Ay(LyzLnxAy(A)))
End Function
Function LyzLnxAy(A() As Lnx) As String()
LyzLnxAy = SyzOyPrp(A, "Lin")
End Function

Function LnxAyeT1Ay(A() As Lnx, T1Ay$()) As Lnx()
Dim L
For Each L In A
    If Not HasEle(T1Ay, T1(CvLnx(L).Lin)) Then PushObj LnxAyeT1Ay, L
Next
End Function

Function LyzLnxAyzWithLno(A() As Lnx) As String()
Dim I, O$()
For Each I In Itr(A)
    With CvLnx(I)
    Push O, FmtQQ("Lno#?:[?]", .Ix, .Lin)
    End With
Next
LyzLnxAyzWithLno = FmtAyzSepSS(O, ":")
End Function

Function ErzLnxAyT1ss(A() As Lnx, T1ss) As String()
Dim T1Ay$(): T1Ay = SySsl(T1ss)
If Si(T1Ay) = 0 Then Exit Function
Dim Er() As Lnx: Er = LnxAyeT1Ay(A, T1Ay)
If Si(Er) = 0 Then Exit Function
ErzLnxAyT1ss = LyzMsgNap("There are lines have invalid T1", "Lines Valid-Ty", LyzLnxAyzWithLno(Er), T1Ay)
End Function

Function LnxAywT2(A() As Lnx, T2) As Lnx()
Dim X
For Each X In Itr(A)
    With CvLnx(X)
        If T2zLin(.Lin) = T2 Then
            PushObj LnxAywT2, X
        End If
    End With
Next
End Function

Function LnxAywRmvT1(A() As Lnx, T) As Lnx()
Dim O()  As Lnx, X
For Each X In Itr(A)
    With CvLnx(X)
        If T1(.Lin) = T Then
            PushObj O, Lnx(.Ix, RmvT1(.Lin))
        End If
    End With
Next
LnxAywRmvT1 = O
End Function

Function LnxRmvT1$(A As Lnx)
If Not IsNothing(A) Then LnxRmvT1 = RmvT1(A.Lin)
End Function

Function LnxStr$(A As Lnx)
LnxStr = "L#" & A.Ix + 1 & ": " & A.Lin
End Function
