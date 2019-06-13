Attribute VB_Name = "QVb_Ay_Op_SyAdd"
Option Compare Text
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

Function TynyzAy(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushI TynyzAy, TypeName(I)
Next
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
AddEle = Ay
Push AddEle, Ele
End Function
Function AddElezStr(Sy$(), Ele$) As String()
Dim O$(): O = Sy: PushI O, Ele: AddElezStr = O
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
Dim Ay1(), Ay2()
GoSub T1
Exit Sub
T1:
    Ay1 = Array(1, 2, 2, 2, 4, 5)
    Ay2 = Array(2, 2)
    Ept = Array(1, 2, 2, 2, 4, 5, 2, 2)
    GoTo Tst
Tst:
    Act = AddAy(Ay1, Ay2)
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

Private Sub Z_AddPfxSfxzAy()
Dim Sy$(), Act$(), Sfx$, Pfx$, Exp$()
Sy = SyzAp(1, 2, 3, 4)
Pfx = "* "
Sfx = "#"
Exp = SyzAp("* 1#", "* 2#", "* 3#", "* 4#")
GoSub Tst
Exit Sub
Tst:
Act = AddPfxSfxzAy(Sy, Pfx, Sfx)
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


Private Sub ZZ()
MVb_AddAy:
End Sub
