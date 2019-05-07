Attribute VB_Name = "QVb_Ay_Op_SyAdd"
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Ay_Op_Add."

Function AddSy(A$(), B$()) As String()
AddSy = A
PushAy AddSy, B
End Function

Function AddAy(A, B)
AddAy = A
PushAy AddAy, B
End Function

Function AddSySorSyAp(Sy$(), ParamArray SorSyAp()) As String()
Dim Av(): Av = SorSyAp
Dim I, J&
For Each I In Av
    If IsSy(I) Then Av(J) = I Else Av(J) = Sy(I)
    J = J + 1
Next
AddSySorSyAp = AddSyAv(Sy, Av)
End Function

Function TyNyzAy(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushI TyNyzAy, TypeName(I)
Next
End Function
Function AddSyAv(Sy$(), SyAv()) As String()
AddSyAv = Sy
Dim I, Ix&
For Each I In Itr(SyAv)
    If Not IsSy(I) Then Thw CSub, "There is Non-Sy in SyAv", "Ix-of-SyAv-Not-Sy TyNy-of-given-SyAv", Ix, TyNyzAy(SyAv)
    PushSy AddSyAv, CvSy(I)
    Ix = Ix + 1
Next
End Function

Function AddSyAp(Sy$(), ParamArray SyAp()) As String()
Dim Av(): Av = SyAp
AddSyAp = AddSyAv(Sy, Av)
End Function

Function AddAyAp(Ay, ParamArray Itm_or_AyAp())
Const CSub$ = CMod & "AddAyAp"
Dim Av(): Av = Itm_or_AyAp
If Not IsArray(Ay) Then Thw CSub, "Fst parameter must be array", "Fst-Pm-TyeName", TypeName(Ay)
Dim I
AddAyAp = Ay
For Each I In Av
    If IsArray(I) Then
        PushAy AddAyAp, I
    Else
        Push AddAyAp, I
    End If
Next
End Function

Function MapAy(Ay, MapFun$) As Variant()
Dim X
For Each X In Itr(Ay)
    Push MapAy, Run(MapFun, X)
Next
End Function

Function AddEle(Ay, Ele)
AddEle = AddEle(Ay, Ele)
End Function
Function AddElezStr(Sy$(), Ele$) As String()
Dim O$(): O = Sy: PushS O, Ele: AddElezStr = O
End Function
Function AddSyItm(Sy$(), Itm$) As String()
AddSyItm = AddAyItm(Sy, Itm)
End Function

Function AddAyItm(Ay, Itm)
Dim O
O = Ay
Push O, Itm
AddAyItm = O
End Function

Function IncAy(Ay, Optional N& = 1)
IncAy = Resi(Ay)
Dim X
For Each X In Itr(Ay)
    PushI IncAy, X + N
Next
End Function

Private Sub Z_AddAy()
Dim Act(), Exp(), Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = AddAy(Ay1, Ay2)
Exp = Array(1, 2, 2, 2, 4, 5, 2, 2)
ThwIfAyabNE Exp, Act
ThwIfAyabNE Ay1, Array(1, 2, 2, 2, 4, 5)
ThwIfAyabNE Ay2, Array(2, 2)
End Sub

Private Sub Z_AddPfxzSy()
Dim Sy$(), Act$(), Pfx$, Exp$()
Sy = SyzAp(1, 2, 3, 4)
Pfx = "* "
Exp = SyzAp("* 1", "* 2", "* 3", "* 4")
GoSub Tst
Exit Sub
Tst:
    Act = AddPfxzSy(Sy, Pfx)
    Debug.Assert IsEqAy(Act, Exp)
    Return
End Sub

Private Sub Z_AddPfxzSySfx()
Dim Sy$(), Act$(), Sfx$, Pfx$, Exp$()
Sy = SyzAp(1, 2, 3, 4)
Pfx = "* "
Sfx = "#"
Exp = SyzAp("* 1#", "* 2#", "* 3#", "* 4#")
GoSub Tst
Exit Sub
Tst:
Act = AddPfxzSySfx(Sy, Pfx, Sfx)
Debug.Assert IsEqAy(Act, Exp)
Return
End Sub

Function TabSy(Sy$()) As String()
TabSy = AddPfxzSy(Sy, vbTab)
End Function

Private Sub Z_AddSySfx()
Dim Sy$(), Act$(), Sfx$, Exp$()
Sy = SyzAp(1, 2, 3, 4)
Sfx = "#"
Exp = SyzAp("1#", "2#", "3#", "4#")
GoSub Tst
Exit Sub
Tst:
    Act = AddSySfx(Sy, Sfx)
    Debug.Assert IsEqAy(Act, Exp)
    Return
End Sub


Private Sub Z()
Z_AddAy
MVb_AddAy:
End Sub
