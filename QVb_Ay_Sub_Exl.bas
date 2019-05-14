Attribute VB_Name = "QVb_Ay_Sub_Exl"
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Ay_Sub_Exl."

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

Function AyeAtCnt(Ay, Optional At = 0, Optional Cnt = 1)
If Cnt <= 0 Then Thw CSub, "Cnt cannot <=0", "At Cnt Ay", At, Cnt, Ay
If Si(Ay) = 0 Then AyeAtCnt = Ay: Exit Function
If At = 0 Then
    If Si(Ay) = Cnt Then
        AyeAtCnt = Resi(Ay)
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

Function AyeFstNEle(Ay, Optional N& = 1)
Dim O: O = Resi(Ay)
Dim J&
For J = N To UB(Ay)
    Push O, Ay(J)
Next
AyeFstNEle = O
End Function


Function AyeEleAt(Ay, Optional At& = 0, Optional Cnt& = 1)
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
Dim O: O = Resi(Ay)
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

Function AyeFmTo(Ay, FmIx, EIx)
Const CSub$ = CMod & "AyeFmTo"
Dim U&
U = UB(Ay)
If 0 > FmIx Or FmIx > U Then Thw CSub, "[FmIx] is out of range", "U FmIx EIx Ay", UB(Ay), FmIx, EIx, Ay
If FmIx > EIx Or EIx > U Then Thw CSub, "[EIx] is out of range", "U FmIx EIx Ay", UB(Ay), FmIx, EIx, Ay
Dim O
    O = Ay
    Dim I&, J&
    I = 0
    For J = EIx + 1 To U
        O(FmIx + I) = O(J)
        I = I + 1
    Next
    Dim Cnt&
    Cnt = EIx - FmIx + 1
    ReDim Preserve O(U - Cnt)
AyeFmTo = O
End Function

Function AyeFstEle(Ay)
AyeFstEle = AyeEleAt(Ay)
End Function

Function AyeFEIx(Ay, B As FEIx)
With B
    AyeFEIx = AyeFmTo(Ay, .FmIx, .EIx)
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

Function AyeIxy(Ay, Ixy)
'Ixy holds index if Ay to be remove.  It has been sorted else will be stop
Ass IsSrtAy(Ay)
Ass IsSrtAy(Ixy)
Dim J&
Dim O: O = Ay
For J = UB(Ixy) To 0 Step -1
    O = AyeEleAt(O, CLng(Ixy(J)))
Next
AyeIxy = O
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
Function PredzLikSy(LikSy$()) As IPred

End Function
Function SyePred(Sy$(), P As IPred) As String()
Dim I
For Each I In Itr(Sy)
    If Not P.Pred(I) Then
        PushI SyePred, CStr(I)
    End If
Next
End Function
Function SyeLikSy(Sy$(), LikSy$()) As String()
SyeLikSy = SyePred(Sy, PredzLikSy(LikSy))
End Function

Function SyeLikssAy(Sy$(), LikssAy$()) As String()
If Si(LikssAy) = 0 Then SyeLikssAy = Sy: Exit Function
SyeLikssAy = SyePred(Sy, PredzLikssAy(LikssAy))
End Function
Function PredzLikssAy(LikssAy$()) As IPred

End Function
Function AyeNegative(Ay)
Dim I
AyeNegative = Resi(Ay)
For Each I In Itr(Ay)
    If I >= 0 Then
        PushI AyeNegative, I
    End If
Next
End Function

Function AyeNEle(Ay, Ele, Cnt%)
If Cnt <= 0 Then Stop
AyeNEle = Resi(Ay)
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

Function SyeT1Sy(Sy$(), ExlT1Sy$()) As String()
'Exclude those Lin in Array-Ay its T1 in ExlT1Ay0
If Si(ExlT1Sy) = 0 Then SyeT1Sy = Sy: Exit Function
SyeT1Sy = SyePred(Sy, PredzInT1Sy(ExlT1Sy))
End Function

Function PredzInT1Sy(T1Ay$()) As IPred
Dim O As PredzInT1Sy
O.Init T1Ay
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

Private Sub Z_AyeFEIx()
Dim Ay
Dim FEIx1 As FEIx
Dim Act
Ay = SplitSpc("a b c d e")
FEIx1 = FEIx(1, 2)
Act = AyeFEIx(Ay, FEIx1)
Ass Si(Act) = 3
Ass JnSpc(Act) = "a d e"
End Sub

Private Sub Z_AyeFEIx1()
Dim Ay
Dim Act
Ay = SplitSpc("a b c d e")
Act = AyeFEIx(Ay, FEIx(1, 2))
Ass Si(Act) = 3
Ass JnSpc(Act) = "a d e"
End Sub

Private Sub Z_AyeIxy()
Dim Ay(), Ixy
Ay = Array("a", "b", "c", "d", "e", "f")
Ixy = Array(1, 3)
Ept = Array("a", "c", "e", "f")
GoSub Tst
Exit Sub
Tst:
    Act = AyeIxy(Ay, Ixy)
    C
    Return
End Sub

Private Sub ZZ()
Z_AyeAtCnt
Z_AyeEmpEleAtEnd
Z_AyeFEIx
Z_AyeFEIx1
Z_AyeIxy
MVb_AySub_Exl:
End Sub

Function RmvBlankzAy(Ay) As String()
Dim I
For Each I In Itr(Ay)
    If Trim(I) <> "" Then
        PushI RmvBlankzAy, I
    End If
Next
End Function


