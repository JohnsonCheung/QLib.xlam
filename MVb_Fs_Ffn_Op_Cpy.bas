Attribute VB_Name = "MVb_Fs_Ffn_Op_Cpy"
Option Explicit
Sub CpyPthzClr(FmPth$, ToPth$)
ThwIfPthNotExist ToPth
ClrPthFil ToPth
Dim Ffn$, I
For Each I In FfnSy(FmPth)
    Ffn = I
    CpyFfnzToPth Ffn, ToPth
Next
End Sub

Sub CpyFfnzUp(Ffn$)
CpyFfnzToPth Ffn, ParPth(Ffn$)
End Sub

Sub CpyFfnSyzToNxt(FfnSy$())
Dim I, Ffn$
For Each I In Itr(FfnSy)
    Ffn = I
    CpyFfnzToNxt Ffn
Next
End Sub

Function CpyFfnzToNxt$(Ffn$)
Dim O$
O = NxtFfn(Ffn$)
CpyFfn Ffn, O
CpyFfnzToNxt = O
End Function
Sub CpyFfnzToPthIfDif(Ffn$, ToPth$, Optional UseEq As Boolean)
If IsEqFfn(Ffn, FfnzPthFn(ToPth, Fn(Ffn)), UseEq) Then Exit Sub
CpyFfnzToPth Ffn, ToPth, OvrWrt:=True
End Sub

Sub CpyFfnSyzIfDif(FfnSy$(), ToPth$, Optional UseEq As Boolean)
Dim I
For Each I In FfnSy
    CpyFfnzIfDif CStr(I), ToPth, UseEq
Next
End Sub

Sub CpyFfn(Ffn$, ToFfn$, Optional OvrWrt As Boolean)
Fso.GetFile(Ffn).Copy ToFfn, OvrWrt
End Sub
Function CpyFfnSy$(FfnSy$(), ToPth$, Optional OvrWrt As Boolean)
Dim Ffn$, I, P$, O$
P = EnsPthSfx(ToPth)
For Each I In FfnSy
    O = P & Fn(Ffn$)
    CpyFfn Ffn, O, OvrWrt
Next

End Function

Function FfnzPthFn$(Pth$, Fn$)
FfnzPthFn = Ffn(Pth, Fn)
End Function

Function Ffn$(Pth$, Fn$)
Ffn = EnsPthSfx(Pth) & Fn
End Function
Function CpyFfnzToPth$(Ffn$, ToPth$, Optional OvrWrt As Boolean)
CpyFfn Ffn, FfnzPthFn(ToPth, Fn(Ffn)), OvrWrt
End Function

Sub CpyFfnzIfDif(Ffn$, ToFfn$, Optional UseEq As Boolean)
If IsEqFfn(Ffn, ToFfn, UseEq) Then
    Dim M$: M = FmtQQ("? file", IIf(UseEq, "Eq", "Same"))
    D LyzFunMsgNap(CSub, M, "FmFfn ToFfn", Ffn, ToFfn)
    Exit Sub
End If
CpyFfn Ffn, ToFfn, OvrWrt:=True
D LyzFunMsgNap(CSub, "File copied", "FmFfn ToFfn", Ffn, ToFfn)
End Sub

Function IsNxtFfn(Ffn$) As Boolean
Dim Las5$: Las5 = Right(Fn(Ffn$), 5)
Select Case True
Case FstChr(Las5) <> "(", LasChr(Las5) <> ")", Not IsDigStr(Mid(Las5, 2, 3))
Case Else: IsNxtFfn = True
End Select
End Function
Function IsDigStr(S) As Boolean
Dim J&
For J = 1 To Len(S)
    If Not IsAscDig(Asc(Mid(S, J, 1))) Then Exit Function
Next
IsDigStr = True
End Function

Function FfnzNxtFfn$(NxtFfn$)
If IsNxtFfn(NxtFfn) Then
    Dim F$: F = Fn(NxtFfn)
    FfnzNxtFfn = Pth(NxtFfn) & Left(F, Len(F) - 5) & Ext(NxtFfn)
Else
    FfnzNxtFfn = NxtFfn
End If
End Function

Function NxtFfn$(Ffn$)
If Not HasFfn(Ffn$) Then NxtFfn = Ffn: Exit Function
Dim J%, O$
For J = 1 To 999
    O = FfnAddFnSfx(Ffn$, "(" & Format(J, "000") & ")")
    If Not HasFfn(O) Then NxtFfn = O: Exit Function
Next
Stop
End Function


