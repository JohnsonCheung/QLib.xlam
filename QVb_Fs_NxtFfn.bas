Attribute VB_Name = "QVb_Fs_NxtFfn"
Option Explicit
Option Compare Text
Private Const CMod$ = "BNxtFfn."
Function NxtNozFfn%(Ffn)
Dim A$: A = Right(RmvExt(Ffn), 5)
If FstChr(A) <> "(" Then Exit Function
If LasChr(A) <> ")" Then Exit Function
Dim M$: M = Mid(A, 2, 3)
If Not IsDigStr(M) Then Exit Function
NxtNozFfn = M
End Function
Function RmvNxtNo$(Ffn)
If IsNxtFfn(Ffn) Then
    Dim A$: A = RmvExt(Ffn)
    RmvNxtNo = RmvLasNChr(A, 5) & Ext(Ffn)
Else
    RmvNxtNo = Ffn
End If
End Function
Private Sub Z_NxtFfn()
Dim Ffn$
'GoSub T0
GoSub T1
Exit Sub
T1: Ffn = "AA(000).xls"
    Ept = "AA(001).xls"
    GoTo Tst
T0:
    Ffn = "AA.xls"
    Ept = "AA(001).xls"
    GoTo Tst
Tst:
    Act = NxtFfn(Ffn)
    C
    Return
End Sub
Function NxtFfn$(Ffn)
Dim J&: J = NxtNozFfn(Ffn)
Dim F$: F = RmvNxtNo(Ffn)
NxtFfn = AddFnSfx(F, "(" & Pad0(J + 1, 3) & ")")
End Function
Function NxtFfnzNotIn(Ffn, NotInFfny$())
Dim J%, O$
O = Ffn
While HasEleS(NotInFfny, O)
    J = J + 1: If J > 1000 Then ThwLoopingTooMuch CSub
    O = NxtFfn(O)
Wend
NxtFfnzNotIn = O
End Function

Function NxtFfnzAva$(Ffn)
Dim J%, O$
O = Ffn
While HasFfn(O)
    If J = 999 Then Thw CSub, "Too much next file in the path of given-ffn", "Given-Ffn", Ffn
    J = J + 1
    O = NxtFfn(O)
Wend
NxtFfnzAva = O
End Function

Function NxtFfny(Ffn) As String() 'Return ffn and all it nxt ffn in the pth of given ffn
If HasFfn(Ffn) Then Push NxtFfny, Ffn  '<==
Dim A$()
    Dim Spec$
        Spec = AddFnSfx(Fn(Ffn), "(???)")
    A = Ffny(Pth(Ffn), Spec)
Dim I, F$
For Each I In Itr(A)
    F = I
    If IsNxtFfn(Ffn) Then PushI NxtFfny, F   '<==
Next
End Function

Function IsNxtFfn(Ffn) As Boolean
Select Case True
Case NxtNozFfn(Ffn) > 0, Right(Fnn(Ffn), 5) = "(000)": IsNxtFfn = True
End Select
End Function


