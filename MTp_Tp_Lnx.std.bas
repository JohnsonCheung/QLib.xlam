Attribute VB_Name = "MTp_Tp_Lnx"
Option Explicit

Function CvLnx(A) As Lnx
Set CvLnx = A
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

Sub LnxBrwAy(A() As Lnx)
BrwAy LnxFmtAySepSS(A)
End Sub

Function LnxFmtAySepSS(A() As Lnx) As String()
Dim I
For Each I In Itr(A)
    With CvLnx(I)
        PushI LnxFmtAySepSS, "L#(" & .Ix & ") " & .Lin
    End With
Next
End Function

Function LyzLnxAy(A() As Lnx) As String()
LyzLnxAy = SyOyP(A, "Lin")
End Function
Function LnxAyeT1Ay(A() As Lnx, T1Ay0) As Lnx()
Dim T1Ay$(), L
T1Ay = CvNy(T1Ay0)
For Each L In A
    If Not HasEle(T1Ay, T1(CvLnx(L).Lin)) Then PushObj LnxAyeT1Ay, L
Next
End Function

Function LnxAyT1Chk(A() As Lnx, T1Ay0) As String()
Dim A1() As Lnx: A1 = LnxAyeT1Ay(A, T1Ay0)
If Sz(A1) = 0 Then Exit Function
Stop
Exit Function
Dim T1Ay$(), mT1$, L, O$()
T1Ay = CvNy(T1Ay0)
For Each L In A
    If Not HasEle(T1Ay, T1(CvLnx(L).Lin)) Then PushI O, L
   
Next
If Sz(O) > 0 Then
    O = AyAddPfx(AyQuoteSq(O), Space(4))
    O = AyInsItm(O, FmtQQ("Following lines have invalid T1.  Valid T1 are [?]", JnSpc(T1Ay)))
End If
LnxAyT1Chk = O
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
