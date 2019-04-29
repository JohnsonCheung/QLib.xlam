Attribute VB_Name = "MVb_Ay_Sub_Exl"
Option Explicit
Const CMod$ = "MVb_Ay_Sub_Exl."

Function AyePatn(Ay, Patn$) As String()
Dim I, Re As New RegExp
Re.Pattern = Patn
For Each I In Itr(Ay)
    If Not Re.Test(I) Then PushI AyePatn, I
Next
End Function
Function AyeRe(Ay, Re As RegExp) As String()
Dim I
For Each I In Itr(Ay)
    If Not Re.Test(I) Then PushI AyeRe, I
Next
End Function
Private Sub Z_AA()
Dim Ay
Ay = Array(1)
Debug.Print VarPtr(Ay)
Debug.Print VarPtr(AA(Ay))
If Not IsEqVar(Ay, AA(Ay)) Then Stop
End Sub
Private Function AA(Ay)
AA = Ay
End Function

Function AyeAtCnt(Ay, Optional At = 0, Optional Cnt = 1)
If Cnt <= 0 Then Thw CSub, "Cnt cannot <=0", "At Cnt Ay", At, Cnt, Ay
If Si(Ay) = 0 Then AyeAtCnt = Ay: Exit Function
If At = 0 Then
    If Si(Ay) = Cnt Then
        AyeAtCnt = AyCln(Ay)
        Exit Function
    End If
End If
Dim U&: U = UB(Ay)
If At > U Then Stop
If At < 0 Then Stop
Dim O: O = Ay
Dim J&
For J = At To U - Cnt
    Asg O(J + Cnt), O(J)
Next
ReDim Preserve O(U - Cnt)
AyeAtCnt = O
End Function
Function PredOfIsDotLin() As IPred

End Function
Function PredOfIsDDLin() As IPred

End Function
Function AyeDDLin(Ay) As String()
AyeDDLin = AywPredFalse(Ay, PredOfIsDDLin)
End Function

Function AyeDotLin(Ay) As String()
AyeDotLin = AywPredFalse(Ay, PredOfIsDotLin)
End Function

Function AyeEle(Ay, Ele) 'Rmv Fst-Ele eq to Ele from Ay
Dim Ix&: Ix = IxzAy(Ay, Ele): If Ix = -1 Then AyeEle = Ay: Exit Function
AyeEle = AyeEleAt(Ay, IxzAy(Ay, Ele))
End Function

Function AyeEleAt(Ay, Optional At = 0, Optional Cnt = 1)
AyeEleAt = AyeAtCnt(Ay, At, Cnt)
End Function

Function AyeEleLik(Ay, Lik$) As String()
If Si(Ay) = 0 Then Exit Function
Dim J&
For J = 0 To UB(Ay)
    If Ay(J) Like Lik Then AyeEleLik = AyeEleAt(Ay, J): Exit Function
Next
End Function

Function AyeEmpEle(Ay)
Dim O: O = AyCln(Ay)
If Si(Ay) > 0 Then
    Dim X
    For Each X In Itr(Ay)
        PushNonEmp O, X
    Next
End If
AyeEmpEle = O
End Function

Function AyeEmpEleAtEnd(Ay)
Dim LasU&, U&
Dim O: O = Ay
For LasU = UB(Ay) To 0 Step -1
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

Function AyeFmTo(Ay, FmIx, ToIx)
Const CSub$ = CMod & "AyeFmTo"
Dim U&
U = UB(Ay)
If 0 > FmIx Or FmIx > U Then Thw CSub, "[FmIx] is out of range", "U FmIx ToIx Ay", UB(Ay), FmIx, ToIx, Ay
If FmIx > ToIx Or ToIx > U Then Thw CSub, "[ToIx] is out of range", "U FmIx ToIx Ay", UB(Ay), FmIx, ToIx, Ay
Dim O
    O = Ay
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

Function AyeFstEle(Ay)
AyeFstEle = AyeEleAt(Ay)
End Function

Function AyeFstNEle(Ay, Optional N = 1)
Dim O: O = Ay: Erase O
Dim J&
For J = N To UB(Ay)
    Push O, Ay(J)
Next
AyeFstNEle = O
End Function

Function AyeFTIx(Ay, B As FTIx)
With B
    If .IsEmp Then AyeFTIx = Ay: Exit Function
    AyeFTIx = AyeFmTo(Ay, .FmIx, .ToIx)
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

Function AyeIxAy(Ay, IxAy)
'IxAy holds index if Ay to be remove.  It has been sorted else will be stop
Ass IsSrtAy(Ay)
Ass IsSrtAy(IxAy)
Dim J&
Dim O: O = Ay
For J = UB(IxAy) To 0 Step -1
    O = AyeEleAt(O, CLng(IxAy(J)))
Next
AyeIxAy = O
End Function

Function AyeLasEle(Ay)
AyeLasEle = AyeEleAt(Ay, UB(Ay))
End Function

Function AyeLasNEle(Ay, Optional NEle% = 1)
If NEle = 0 Then AyeLasNEle = Ay: Exit Function
Dim O: O = Ay
Select Case Si(Ay)
Case Is > NEle:    ReDim Preserve O(UB(Ay) - NEle)
Case NEle: Erase O
Case Else: Stop
End Select
AyeLasNEle = O
End Function
Function PredzLik(Lik$) As IPred

End Function
Function SyeLik(Sy$(), Lik$) As String()
SyeLik = SyePred(Sy, PredzLik(Lik))
End Function
Function PredzLikAy(LikAy$()) As IPred

End Function
Function SyePred(Sy$(), P As IPred) As String()
Dim I
For Each I In Itr(Sy)
    If Not P.Pred(I) Then
        PushS SyePred, CStr(I)
    End If
Next
End Function
Function SyeLikAy(Sy$(), LikAy$()) As String()
SyeLikAy = SyePred(Sy, PredzLikAy(LikAy))
End Function

Function SyeLikssAy(Sy$(), LikssAy$()) As String()
If Si(LikssAy) = 0 Then SyeLikssAy = Sy: Exit Function
SyeLikssAy = SyePred(Sy, PredzLikssAy(LikssAy))
End Function
Function PredzLikssAy(LikssAy$()) As IPred

End Function
Function AyeNegative(Ay)
Dim I
AyeNegative = AyCln(Ay)
For Each I In Itr(Ay)
    If I >= 0 Then
        PushI AyeNegative, I
    End If
Next
End Function

Function AyeNEle(Ay, Ele, Cnt%)
If Cnt <= 0 Then Stop
AyeNEle = AyCln(Ay)
Dim X, C%
C = Cnt
For Each X In Itr(Ay)
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
Function PredzIsOneTermLin() As IPred

End Function
Function SyeOneTermLin(Sy$()) As String()
SyeOneTermLin = SyePred(Sy, PredzIsOneTermLin)
End Function
Function PredzPfx(Pfx$) As IPred

End Function
Function SyePfx(Sy$(), ExlPfx$) As String()
SyePfx = SyePred(Sy, PredzPfx(ExlPfx))
End Function

Function SyeT1Ay(Sy$(), ExlT1Sy$()) As String()
'Exclude those Lin in Array-Ay its T1 in ExlT1Ay0
If Si(ExlT1Sy) = 0 Then SyeT1Ay = Sy: Exit Function
SyeT1Ay = SyePred(Sy, PredzInT1Sy(ExlT1Sy))
End Function

Function PredzInT1Sy(T1Sy$()) As IPred
Dim O As PredzInT1Sy
O.Init T1Sy
Set PredzInT1Sy = O
End Function

Private Sub Z_AyeAtCnt()
Dim Ay()
Ay = Array(1, 2, 3, 4, 5)
Ept = Array(1, 4, 5)
GoSub Tst
'
Exit Sub

Tst:
    Act = AyeAtCnt(Ay, 1, 2)
    C
    Return
End Sub

Private Sub Z_AyeEmpEleAtEnd()
Dim Ay: Ay = Array(Empty, Empty, Empty, 1, Empty, Empty)
Dim Act: Act = AyeEmpEleAtEnd(Ay)
Ass Si(Act) = 4
Ass Act(3) = 1
End Sub

Private Sub Z_AyeFTIx()
Dim Ay
Dim FTIx1 As FTIx
Dim Act
Ay = SplitSpc("a b c d e")
Set FTIx1 = FTIx(1, 2)
Act = AyeFTIx(Ay, FTIx1)
Ass Si(Act) = 3
Ass JnSpc(Act) = "a d e"
End Sub

Private Sub Z_AyeFTIx1()
Dim Ay
Dim Act
Ay = SplitSpc("a b c d e")
Act = AyeFTIx(Ay, FTIx(1, 2))
Ass Si(Act) = 3
Ass JnSpc(Act) = "a d e"
End Sub

Private Sub Z_AyeIxAy()
Dim Ay(), IxAy
Ay = Array("a", "b", "c", "d", "e", "f")
IxAy = Array(1, 3)
Ept = Array("a", "c", "e", "f")
GoSub Tst
Exit Sub
Tst:
    Act = AyeIxAy(Ay, IxAy)
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


