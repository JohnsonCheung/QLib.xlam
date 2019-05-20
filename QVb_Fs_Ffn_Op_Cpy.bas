Attribute VB_Name = "QVb_Fs_Ffn_Op_Cpy"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Fs_Ffn_Op_Cpy."
Private Const Asm$ = "QVb"
Sub CpyPthzClr(FmPth$, ToPth$)
ThwIf_PthNotExist ToPth, CSub
ClrPthFil ToPth
Dim Ffn$, I
For Each I In Ffny(FmPth)
    Ffn = I
    CpyFfnzToPth Ffn, ToPth
Next
End Sub

Sub CpyFfnzUp(Ffn)
CpyFfnzToPth Ffn, ParPth(Ffn)
End Sub

Sub CpyFfnyzToNxt(Ffny$())
Dim I, Ffn$
For Each I In Itr(Ffny)
    Ffn = I
    CpyFfnzToNxt Ffn
Next
End Sub

Function CpyFfnzToNxt$(Ffn)
Dim O$
O = NxtFfnzAva(Ffn)
CpyFfn Ffn, O
CpyFfnzToNxt = O
End Function

Sub CpyFfnzToPthIfDif(Ffn, ToPth$, Optional B As EmFilCmp = EmFilCmp.EiCmpEq)
Dim Fn$: Fn = FnzFfn(Ffn)
Dim ToFfn$: ToFfn = FfnzPthFn(ToPth, Fn)
If IsEqFfn(Ffn, ToFfn, B) Then Exit Sub
CpyFfnzToPth Ffn, ToPth, OvrWrt:=True
End Sub

Sub CpyFfnyzIfDif(Ffny$(), ToPth$, Optional B As EmFilCmp = EmFilCmp.EiCmpEq)
Dim I
For Each I In Ffny
    CpyFfnzIfDif CStr(I), ToPth, B
Next
End Sub

Sub CpyFfn(Ffn, ToFfn$, Optional OvrWrt As Boolean)
Fso.GetFile(Ffn).Copy ToFfn, OvrWrt
End Sub

Function CpyFfny$(Ffny$(), ToPth$, Optional OvrWrt As Boolean)
Dim Ffn$, I, P$, O$
P = EnsPthSfx(ToPth)
For Each I In Ffny
    O = P & Fn(Ffn)
    CpyFfn Ffn, O, OvrWrt
Next

End Function

Function FfnzPthFn$(Pth, Fn$)
FfnzPthFn = Ffn(Pth, Fn)
End Function

Function Ffn$(Pth, Fn$)
Ffn = EnsPthSfx(Pth) & Fn
End Function
Function CpyFfnzToPth$(Ffn, ToPth$, Optional OvrWrt As Boolean)
CpyFfn Ffn, FfnzPthFn(ToPth, Fn(Ffn)), OvrWrt
End Function

Sub CpyFfnzIfDif(Ffn, ToFfn$, Optional B As EmFilCmp)
If IsEqFfn(Ffn, ToFfn, B) Then
    Dim M$: M = FmtQQ("? file", IIf(B = EiCmpEq, "Eq", "Same"))
    D LyzFunMsgNap(CSub, M, "FmFfn ToFfn", Ffn, ToFfn)
    Exit Sub
End If
CpyFfn Ffn, ToFfn, OvrWrt:=True
D LyzFunMsgNap(CSub, "File copied", "FmFfn ToFfn", Ffn, ToFfn)
End Sub

Function IsDigStr(S) As Boolean
Dim J&
For J = 1 To Len(S)
    If Not IsAscDig(Asc(Mid(S, J, 1))) Then Exit Function
Next
IsDigStr = True
End Function


