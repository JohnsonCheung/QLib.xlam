Attribute VB_Name = "MVb_Ay_Sub_Exl"
Option Explicit
Const CMod$ = "MVb_Ay_Sub_Exl."

Function AyePatn(A, Patn$) As String()
Dim I, Re As New RegExp
Re.Pattern = Patn
For Each I In Itr(A)
    If Not Re.Test(I) Then PushI AyePatn, I
Next
End Function
Function AyeRe(Ay, Re As RegExp) As String()
Dim I
For Each I In Itr(Ay)
    If Not Re.Test(I) Then PushI AyeRe, I
Next
End Function
Sub Z_AA()
Dim A
A = Array(1)
Debug.Print VarPtr(A)
Debug.Print VarPtr(AA(A))
If Not IsEqVar(A, AA(A)) Then Stop
End Sub
Function AA(A)
AA = A
End Function
Function AyeAtCnt(A, Optional At = 0, Optional Cnt = 1)
If Cnt <= 0 Then Thw CSub, "Cnt cannot <=0", "At Cnt Ay", At, Cnt, A
If Si(A) = 0 Then AyeAtCnt = A: Exit Function
If At = 0 Then
    If Si(A) = Cnt Then
        AyeAtCnt = AyCln(A)
        Exit Function
    End If
End If
Dim U&: U = UB(A)
If At > U Then Stop
If At < 0 Then Stop
Dim O: O = A
Dim J&
For J = At To U - Cnt
    Asg O(J + Cnt), O(J)
Next
ReDim Preserve O(U - Cnt)
AyeAtCnt = O
End Function

Function AyeDDLin(A) As String()
AyeDDLin = AywPredFalse(A, "IsDDLin")
End Function

Function AyeDotLin(A) As String()
AyeDotLin = AywPredFalse(A, "IsDotLin")
End Function

Function AyeEle(A, Ele) 'Rmv Fst-Ele eq to Ele from Ay
Dim Ix&: Ix = IxzAy(A, Ele): If Ix = -1 Then AyeEle = A: Exit Function
AyeEle = AyeEleAt(A, IxzAy(A, Ele))
End Function

Function AyeEleAt(Ay, Optional At = 0, Optional Cnt = 1)
AyeEleAt = AyeAtCnt(Ay, At, Cnt)
End Function

Function AyeEleLik(A, Lik$) As String()
If Si(A) = 0 Then Exit Function
Dim J&
For J = 0 To UB(A)
    If A(J) Like Lik Then AyeEleLik = AyeEleAt(A, J): Exit Function
Next
End Function

Function AyeEmpEle(A)
Dim O: O = AyCln(A)
If Si(A) > 0 Then
    Dim X
    For Each X In Itr(A)
        PushNonEmp O, X
    Next
End If
AyeEmpEle = O
End Function

Function AyeEmpEleAtEnd(A)
Dim LasU&, U&
Dim O: O = A
For LasU = UB(A) To 0 Step -1
    If Not IsEmp(O(LasU)) Then
        Exit For
    End If
Next
If LasU = -1 Then
    Erase O
Else
    ReDim Preserve O(LasU)
End If
AyeEmpEleAtEnd = O
End Function

Function AyeFmTo(A, FmIx, ToIx)
Const CSub$ = CMod & "AyeFmTo"
Dim U&
U = UB(A)
If 0 > FmIx Or FmIx > U Then Thw CSub, "[FmIx] is out of range", "U FmIx ToIx Ay", UB(A), FmIx, ToIx, A
If FmIx > ToIx Or ToIx > U Then Thw CSub, "[ToIx] is out of range", "U FmIx ToIx Ay", UB(A), FmIx, ToIx, A
Dim O
    O = A
    Dim I&, J&
    I = 0
    For J = ToIx + 1 To U
        O(FmIx + I) = O(J)
        I = I + 1
    Next
    Dim Cnt&
    Cnt = ToIx - FmIx + 1
    ReDim Preserve O(U - Cnt)
AyeFmTo = O
End Function

Function AyeFstEle(A)
AyeFstEle = AyeEleAt(A)
End Function

Function AyeFstNEle(A, Optional N = 1)
Dim O: O = A: Erase O
Dim J&
For J = N To UB(A)
    Push O, A(J)
Next
AyeFstNEle = O
End Function

Function AyeFTIx(A, B As FTIx)
With B
    If .IsEmp Then AyeFTIx = A: Exit Function
    AyeFTIx = AyeFmTo(A, .FmIx, .ToIx)
End With
End Function

Function AyeIxSet(Ay, IxSet As Aset)
Dim J&, O
O = Ay: Erase O
For J = 0 To UBound(Ay)
    If Not IxSet.Has(J) Then PushI O, Ay(J)
Next
AyeIxSet = O
End Function

Function AyeIxAy(A, IxAy)
'IxAy holds index if A to be remove.  It has been sorted else will be stop
Ass IsSrtAy(A)
Ass IsSrtAy(IxAy)
Dim J&
Dim O: O = A
For J = UB(IxAy) To 0 Step -1
    O = AyeEleAt(O, CLng(IxAy(J)))
Next
AyeIxAy = O
End Function

Function AyeLasEle(A)
AyeLasEle = AyeEleAt(A, UB(A))
End Function

Function AyeLasNEle(A, Optional NEle% = 1)
If NEle = 0 Then AyeLasNEle = A: Exit Function
Dim O: O = A
Select Case Si(A)
Case Is > NEle:    ReDim Preserve O(UB(A) - NEle)
Case NEle: Erase O
Case Else: Stop
End Select
AyeLasNEle = O
End Function

Function AyeLik(A, Lik) As String()
Dim I
For Each I In Itr(A)
    If Not I Like Lik Then PushI AyeLik, I
Next
End Function

Function AyeLikAy(A, LikeAy$()) As String()
Dim I
For Each I In Itr(A)
    If Not HitLikAy(I, LikeAy) Then Push AyeLikAy, I
Next
End Function

Function AyeLikss(A, Likss$) As String()
AyeLikss = AyeLikAy(A, SySsl(Likss))
End Function

Function AyeLikssAy(A, LikssAy$()) As String()
If Si(LikssAy) = 0 Then AyeLikssAy = SyzAy(A): Exit Function
Dim Likss
For Each Likss In Itr(A)
    If Not HitLikss(A, Likss) Then PushI AyeLikssAy, A
Next
End Function

Function AyeNeg(A)
Dim I
AyeNeg = AyCln(A)
For Each I In Itr(A)
    If I >= 0 Then
        PushI AyeNeg, I
    End If
Next
End Function

Function AyeNEle(A, Ele, Cnt%)
If Cnt <= 0 Then Stop
AyeNEle = AyCln(A)
Dim X, C%
C = Cnt
For Each X In Itr(A)
    If C = 0 Then
        PushI AyeNEle, X
    Else
        If X <> Ele Then
            Push AyeNEle, X
        Else
            C = C - 1
        End If
    End If
Next
X:
End Function

Function AyeOneTermLin(A) As String()
AyeOneTermLin = AywPredFalse(A, "LinIsOneTermLin")
End Function

Function AyePfx(A, ExlPfx$) As String()
Dim I
For Each I In Itr(A)
    If Not HasPfx(I, ExlPfx) Then PushI AyePfx, I
Next
End Function

Function AyeT1Ay(A, ExlT1Ay0) As String()
'Exclude those Lin in Array-A its T1 in ExlT1Ay0
Dim Exl$(): Exl = NyzNN(ExlT1Ay0): If Si(Exl) = 0 Then Stop
Dim L
For Each L In Itr(A)
    If Not HasEle(Exl, T1(L)) Then
        PushI AyeT1Ay, L
    End If
Next
End Function


Private Sub Z_AyeAtCnt()
Dim A()
A = Array(1, 2, 3, 4, 5)
Ept = Array(1, 4, 5)
GoSub Tst
'
Exit Sub

Tst:
    Act = AyeAtCnt(A, 1, 2)
    C
    Return
End Sub

Private Sub Z_AyeEmpEleAtEnd()
Dim A: A = Array(Empty, Empty, Empty, 1, Empty, Empty)
Dim Act: Act = AyeEmpEleAtEnd(A)
Ass Si(Act) = 4
Ass Act(3) = 1
End Sub

Private Sub Z_AyeFTIx()
Dim A
Dim FTIx1 As FTIx
Dim Act
A = SplitSpc("a b c d e")
Set FTIx1 = FTIx(1, 2)
Act = AyeFTIx(A, FTIx1)
Ass Si(Act) = 3
Ass JnSpc(Act) = "a d e"
End Sub

Private Sub Z_AyeFTIx1()
Dim A
Dim Act
A = SplitSpc("a b c d e")
Act = AyeFTIx(A, FTIx(1, 2))
Ass Si(Act) = 3
Ass JnSpc(Act) = "a d e"
End Sub

Private Sub Z_AyeIxAy()
Dim A(), IxAy
A = Array("a", "b", "c", "d", "e", "f")
IxAy = Array(1, 3)
Ept = Array("a", "c", "e", "f")
GoSub Tst
Exit Sub
Tst:
    Act = AyeIxAy(A, IxAy)
    C
    Return
End Sub

Private Sub Z()
Z_AyeAtCnt
Z_AyeEmpEleAtEnd
Z_AyeFTIx
Z_AyeFTIx1
Z_AyeIxAy
MVb_AySub_Exl:
End Sub

Function SyRmvBlank(Ay$()) As String()
Dim I
For Each I In Itr(Ay)
    If Trim(I) <> "" Then
        PushI SyRmvBlank, I
    End If
Next
End Function


