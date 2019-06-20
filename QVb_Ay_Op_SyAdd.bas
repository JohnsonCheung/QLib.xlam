Attribute VB_Name = "QVb_Ay_Op_SyAdd"
Option Compare Text
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Ay_Op_Add."

Function SyzAdd(A$(), B$()) As String()
SyzAdd = A
PushIAy SyzAdd, B
End Function

Function AyzAdd(AyA, AyB)
AyzAdd = AyA
PushAy AyzAdd, AyB
End Function

Function VbTyNyzAy(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushI VbTyNyzAy, TypeName(I)
Next
End Function

Function AyzAddAp(Ay, ParamArray Itm_or_AyAp())
Const CSub$ = CMod & "AyzAddAp"
Dim Av(): Av = Itm_or_AyAp
If Not IsArray(Ay) Then Thw CSub, "Fst parameter must be array", "Fst-Pm-TyeName", TypeName(Ay)
Dim I
AyzAddAp = Ay
For Each I In Av
    If IsArray(I) Then
        PushAy AyzAddAp, I
    Else
        Push AyzAddAp, I
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
AddEle = Ay
Push AddEle, Ele
End Function
Function AddElezStr(Sy$(), Ele$) As String()
Dim O$(): O = Sy: PushI O, Ele: AddElezStr = O
End Function
Function SyzAddItm(Sy$(), Itm$) As String()
SyzAddItm = AyzAddItm(Sy, Itm)
End Function

Function AyzAddItm(Ay, Itm)
Dim O
O = Ay
Push O, Itm
AyzAddItm = O
End Function

Function IncAy(Ay, Optional N& = 1)
IncAy = ResiU(Ay)
Dim X
For Each X In Itr(Ay)
    PushI IncAy, X + N
Next
End Function

Private Sub Z_AyzAdd()
Dim Ay1(), Ay2()
GoSub T1
Exit Sub
T1:
    Ay1 = Array(1, 2, 2, 2, 4, 5)
    Ay2 = Array(2, 2)
    Ept = Array(1, 2, 2, 2, 4, 5, 2, 2)
    GoTo Tst
Tst:
    Act = AyzAdd(Ay1, Ay2)
    C
    Return
End Sub

Private Sub Z_AddPfxzAy()
Dim Sy$(), Pfx$
GoSub T1
Exit Sub
T1:
    Sy = SyzSS("1 2 3 4")
    Pfx = "* "
    Ept = SyzAp("* 1", "* 2", "* 3", "* 4")
    GoTo Tst
Tst:
    Act = AddPfxzAy(Sy, Pfx)
    C
    Return
End Sub

Private Sub Z_AddPfxSzAy()
Dim Sy$(), Act$(), Sfx$, Pfx$, Exp$()
Sy = SyzAp(1, 2, 3, 4)
Pfx = "* "
Sfx = "#"
Exp = SyzAp("* 1#", "* 2#", "* 3#", "* 4#")
GoSub Tst
Exit Sub
Tst:
Act = AddPfxSzAy(Sy, Pfx, Sfx)
Debug.Assert IsEqAy(Act, Exp)
Return
End Sub

Function TabSy(Sy$()) As String()
TabSy = AddPfxzAy(Sy, vbTab)
End Function

Private Sub Z_AddSfxzAy()
Dim Sy$(), Sfx$
Sy = SyzSS("1 2 3 4")
Sfx = "#"
Ept = SyzSS("1# 2# 3# 4#")
GoSub Tst
Exit Sub
Tst:
    Act = AddSfxzAy(Sy, Sfx)
    C
    Return
End Sub


Private Sub Z()
MVb_AyzAdd:
End Sub
