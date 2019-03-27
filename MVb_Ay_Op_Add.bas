Attribute VB_Name = "MVb_Ay_Op_Add"
Option Explicit
Const CMod$ = "MVb_Ay_Add."

Function AyAdd1(A)
AyAdd1 = AyAddN(A, 1)
End Function

Function SyAdd(A$(), B$()) As String()
SyAdd = A
PushAy SyAdd, B
End Function

Function AyAdd(A, B)
AyAdd = A
PushAy AyAdd, B
End Function

Function SyAddSorSyAp(GivenSy$(), ParamArray SorSyAp()) As String()
Dim Av(): Av = SorSyAp
Dim I, J&
For Each I In Av
    If IsSy(I) Then Av(J) = I Else Av(J) = Sy(I)
    J = J + 1
Next
SyAddSorSyAp = SyAddSyAv(GivenSy, Av)
End Function
Function TyNyzAy(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushI TyNyzAy, TypeName(I)
Next
End Function
Function SyAddSyAv(Sy$(), SyAv()) As String()
SyAddSyAv = Sy
Dim I
For Each I In Itr(SyAv)
    If Not IsSy(I) Then Thw CSub, "There is Non-Sy in SyAv", "TyNy-of-given-SyAv", TyNyzAy(SyAv)
    PushIAy SyAddSyAv, I
Next
End Function

Function SyAddAp(Sy$(), ParamArray SyAp()) As String()
Dim Av(): Av = SyAp
SyAddAp = SyAddSyAv(Sy, Av)
End Function

Function AyAddAp(Ay, ParamArray Itm_or_AyAp())
Const CSub$ = CMod & "AyAddAp"
Dim Av(): Av = Itm_or_AyAp
If Not IsArray(Ay) Then Thw CSub, "Fst parameter must be array", "Fst-Pm-TyeName", TypeName(Ay)
Dim I
AyAddAp = Ay
For Each I In Av
    If IsArray(I) Then
        PushAy AyAddAp, I
    Else
        Push AyAddAp, I
    End If
Next
End Function
Function AyMap(Ay, Map$) As Variant()
Dim X
For Each X In Itr(Ay)
    Push AyMap, Run(Map, X)
Next
End Function

Function DryzAyMap(Ay, Map$) As Variant()
Dim X
For Each X In Itr(Ay)
    PushI DryzAyMap, Array(X, Run(Map, X))
Next
End Function

Function AyAddItm(A, Itm)
Dim O
O = A
Push O, Itm
AyAddItm = O
End Function

Function AyAddN(A, N)
AyAddN = AyCln(A)
Dim X
For Each X In Itr(A)
    PushI AyAddN, X + N
Next
End Function

Private Sub Z_AyAdd()
Dim Act(), Exp(), Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = AyAdd(Ay1, Ay2)
Exp = Array(1, 2, 2, 2, 4, 5, 2, 2)
ThwAyabNE Exp, Act
ThwAyabNE Ay1, Array(1, 2, 2, 2, 4, 5)
ThwAyabNE Ay2, Array(2, 2)
End Sub


Private Sub ZZ_AyAdd()
Dim Act(), Exp(), Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = AyAdd(Ay1, Ay2)
Exp = Array(1, 2, 2, 2, 4, 5, 2, 2)
ThwIfNE Exp, Act
ThwIfNE Ay1, Array(1, 2, 2, 2, 4, 5)
ThwIfNE Ay2, Array(2, 2)
End Sub

Private Sub ZZ_AyAddPfx()
Dim A, Act$(), Pfx$, Exp$()
A = Array(1, 2, 3, 4)
Pfx = "* "
Exp = Sy("* 1", "* 2", "* 3", "* 4")
GoSub Tst
Exit Sub
Tst:
Act = AyAddPfx(A, Pfx)
Debug.Assert IsEqAy(Act, Exp)
Return
End Sub

Private Sub ZZ_AyAddPfxSfx()
Dim A, Act$(), Sfx$, Pfx$, Exp$()
A = Array(1, 2, 3, 4)
Pfx = "* "
Sfx = "#"
Exp = Sy("* 1#", "* 2#", "* 3#", "* 4#")
GoSub Tst
Exit Sub
Tst:
Act = AyAddPfxSfx(A, Pfx, Sfx)
Debug.Assert IsEqAy(Act, Exp)
Return
End Sub

Function AyTab(A) As String()
AyTab = AyAddPfx(A, vbTab)
End Function

Private Sub ZZ_AyAddSfx()
Dim A, Act$(), Sfx$, Exp$()
A = Array(1, 2, 3, 4)
Sfx = "#"
Exp = Sy("1#", "2#", "3#", "4#")
GoSub Tst
Exit Sub
Tst:
Act = AyAddSfx(A, Sfx)
Debug.Assert IsEqAy(Act, Exp)
Return
End Sub


Private Sub Z()
Z_AyAdd
MVb_AyAdd:
End Sub
