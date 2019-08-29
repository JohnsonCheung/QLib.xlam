Attribute VB_Name = "QVb_Ay_SubSet_Ae"
Option Compare Text
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Ay_Sub_Exl."

Function AePatn(Ay, Patn$) As String()
Dim I, Re As New RegExp
Re.Pattern = Patn
For Each I In Itr(Ay)
    If Not Re.Test(I) Then PushI AePatn, I
Next
End Function
Function AeRe(Ay, Re As RegExp) As String()
Dim I
For Each I In Itr(Ay)
    If Not Re.Test(I) Then PushI AeRe, I
Next
End Function
Function AeVbRmk(Ay) As String()
Dim L: For Each L In Itr(Ay)
    If Not IsLinVbRmk(L) Then PushI AeVbRmk, L
Next
End Function
Function AwVbRmk(Ay) As String()
Dim L: For Each L In Itr(Ay)
    If IsLinVbRmk(L) Then PushI AwVbRmk, L
Next
End Function
Function AeAtCnt(Ay, Optional At = 0, Optional Cnt = 1)
If Cnt <= 0 Then Thw CSub, "Cnt cannot <=0", "At Cnt Ay", At, Cnt, Ay
If Si(Ay) = 0 Then AeAtCnt = Ay: Exit Function
If At = 0 Then
    If Si(Ay) = Cnt Then
        AeAtCnt = ResiU(Ay)
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
AeAtCnt = O
End Function

Function AeBlnk(Ay) As String()
Dim I: For Each I In Itr(Ay)
    If Trim(I) <> "" Then PushI AeBlnk, I
Next
End Function


Function AeEle(Ay, Ele) 'Rmv Fst-Ele eq to Ele from Ay
Dim Ix&: Ix = IxzAy(Ay, Ele): If Ix = -1 Then AeEle = Ay: Exit Function
AeEle = AeEleAt(Ay, IxzAy(Ay, Ele))
End Function

Function AeFstNEle(Ay, Optional N& = 1)
Dim O: O = ResiU(Ay)
Dim J&
For J = N To UB(Ay)
    Push O, Ay(J)
Next
AeFstNEle = O
End Function

Function AeEleAt(Ay, Optional At& = 0, Optional Cnt& = 1)
AeEleAt = AeAtCnt(Ay, At, Cnt)
End Function

Function AeEleLik(Ay, Lik$) As String()
If Si(Ay) = 0 Then Exit Function
Dim J&
For J = 0 To UB(Ay)
    If Ay(J) Like Lik Then AeEleLik = AeEleAt(Ay, J): Exit Function
Next
End Function

Function AeEmpEle(Ay)
Dim O: O = ResiU(Ay)
If Si(Ay) > 0 Then
    Dim X
    For Each X In Itr(Ay)
        PushNonEmp O, X
    Next
End If
AeEmpEle = O
End Function

Function AeBlnkStr(Ay) As String()
Dim X
For Each X In Itr(Ay)
    If Trim(X) <> "" Then
        PushI AeBlnkStr, X
    End If
Next
End Function

Function AeEmpEleAtEnd(Ay)
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
AeEmpEleAtEnd = O
End Function

Function AeFmTo(Ay, FmIx, EIx)
Const CSub$ = CMod & "AeFmTo"
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
AeFmTo = O
End Function

Function AeFstLas(Ay)
Dim J&
AeFstLas = Ay
Erase AeFstLas
For J = 1 To UB(Ay) - 1
    PushI AeFstLas, Ay(J)
Next
End Function

Function AeFstEle(Ay)
AeFstEle = AeEleAt(Ay)
End Function

Function AeFei(Ay, B As Fei)
With B
    AeFei = AeFmTo(Ay, .FmIx, .EIx)
End With
End Function

Function AeIxSet(Ay, IxSet As Aset)
Dim J&, O
O = Ay: Erase O
For J = 0 To UBound(Ay)
    If Not IxSet.Has(J) Then PushI O, Ay(J)
Next
AeIxSet = O
End Function

Function AeIxy(Ay, SrtdIxy)
'Fm SrtdIxy : holds index if Ay to be remove.  It has been sorted else will be stop
Ass IsArray(Ay)
Ass IsSrtedzAy(SrtdIxy)
Dim J&
Dim O: O = Ay
For J = UB(SrtdIxy) To 0 Step -1
    O = AeEleAt(O, CLng(SrtdIxy(J)))
Next
AeIxy = O
End Function

Function AeLasEle(Ay)
AeLasEle = AeEleAt(Ay, UB(Ay))
End Function

Function AeLasNEle(Ay, Optional NEle% = 1)
If NEle = 0 Then AeLasNEle = Ay: Exit Function
Dim O: O = Ay
Select Case Si(Ay)
Case Is > NEle:    ReDim Preserve O(UB(Ay) - NEle)
Case NEle: Erase O
Case Else: Stop
End Select
AeLasNEle = O
End Function
Function PredzLik(Lik$) As IPred

End Function
Function SyeLik(Sy$(), Lik$) As String()
SyeLik = SyePred(Sy, PredzLik(Lik))
End Function

Function PredzLikAy(LikAy$()) As PredLikAy
Set PredzLikAy = New PredLikAy
PredzLikAy.Init LikAy
End Function

Function SyePred(Sy$(), P As IPred) As String()
Dim I
For Each I In Itr(Sy)
    If Not P.Pred(I) Then
        PushI SyePred, I
    End If
Next
End Function

Function SyeLikAy(Sy$(), LikAy$()) As String()
SyeLikAy = SyePred(Sy, PredzLikAy(LikAy))
End Function

Function SyeKssAy(Sy$(), KssAy$()) As String()
If Si(KssAy) = 0 Then SyeKssAy = Sy: Exit Function
SyeKssAy = SyePred(Sy, PredzKssAy(KssAy))
End Function

Function PredzKssAy(KssAy$()) As PredLikAy
Dim O As New PredLikAy
O.Init KssAy
Set PredzKssAy = O
End Function

Private Sub Z_SyeKss()
Dim Sy$(), Kss$
GoSub Z
GoSub T0
Exit Sub
T0:
    Sy = SyzSS("A B C CD E E1 E3")
    Kss = "C* E*"
    Ept = SyzSS("A B")
    GoTo Tst
Z:
    D SyeKss(SyzSS("A B C CD E E1 E3"), "C* E*")
    Return
Tst:
    Act = SyeKss(Sy, Kss)
    C
    Return
End Sub

Function SyeKss(Sy$(), Kss$) As String()
If Kss = "" Then SyeKss = Sy: Exit Function
SyeKss = SyePred(Sy, PredzLikAy(SyzSS(Kss)))
End Function

Function AeNegative(Ay)
Dim I
AeNegative = ResiU(Ay)
For Each I In Itr(Ay)
    If I >= 0 Then
        PushI AeNegative, I
    End If
Next
End Function

Function AeNEle(Ay, Ele, Cnt%)
If Cnt <= 0 Then Stop
AeNEle = ResiU(Ay)
Dim X, C%
C = Cnt
For Each X In Itr(Ay)
    If C = 0 Then
        PushI AeNEle, X
    Else
        If X <> Ele Then
            Push AeNEle, X
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
Function PredzPfx(Pfx) As IPred
Dim O As New PredPfx
O.Init Pfx
Set PredzPfx = O
End Function
Function PredzSubStr(SubStr) As IPred
Dim O As New PredSubStr
O.Init SubStr
Set PredzSubStr = O
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
Dim O As PredInT1Sy
O.Init T1Ay
Set PredzInT1Sy = O
End Function

Private Sub Z_AeAtCnt()
Dim Ay()
Ay = Array(1, 2, 3, 4, 5)
Ept = Array(1, 4, 5)
GoSub Tst
'
Exit Sub

Tst:
    Act = AeAtCnt(Ay, 1, 2)
    C
    Return
End Sub

Private Sub Z_AeEmpEleAtEnd()
Dim Ay: Ay = Array(Empty, Empty, Empty, 1, Empty, Empty)
Dim Act: Act = AeEmpEleAtEnd(Ay)
Ass Si(Act) = 4
Ass Act(3) = 1
End Sub

Private Sub Z_AeFei()
Dim Ay
Dim Fei1 As Fei
Dim Act
Ay = SplitSpc("a b c d e")
Fei1 = Fei(1, 2)
Act = AeFei(Ay, Fei1)
Ass Si(Act) = 3
Ass JnSpc(Act) = "a d e"
End Sub

Private Sub Z_AeFei1()
Dim Ay
Dim Act
Ay = SplitSpc("a b c d e")
Act = AeFei(Ay, Fei(1, 2))
Ass Si(Act) = 3
Ass JnSpc(Act) = "a d e"
End Sub

Private Sub Z_AeIxy()
Dim Ay(), Ixy
Ay = Array("a", "b", "c", "d", "e", "f")
Ixy = Array(1, 3)
Ept = Array("a", "c", "e", "f")
GoSub Tst
Exit Sub
Tst:
    Act = AeIxy(Ay, Ixy)
    C
    Return
End Sub

Private Sub Z()
Z_AeAtCnt
Z_AeEmpEleAtEnd
Z_AeFei
Z_AeFei1
Z_AeIxy
MVb_AySub_Exl:
End Sub

Function RmvBlnkzAy(Ay) As String()
Dim I
For Each I In Itr(Ay)
    If Trim(I) <> "" Then
        PushI RmvBlnkzAy, I
    End If
Next
End Function


Function AeKss(Ay, ExlKss$) As String()
Stop
'AeKss = AePred(Ay, PredzIsKss(ExlKss))
End Function

Function AePfx(Ay, Pfx$) As String()
Dim I
For Each I In Itr(Ay)
    If Not HasPfx(I, Pfx) Then PushI AePfx, I
Next
End Function

Function AePred(Ay, P As IPred) As String()
Dim I
For Each I In Itr(Ay)
    If Not P.Pred(I) Then
        PushI AePred, I
    End If
Next
End Function



'
