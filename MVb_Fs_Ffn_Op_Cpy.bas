Attribute VB_Name = "MVb_Fs_Ffn_Op_Cpy"
Option Explicit
Sub CpyPthzClr(FmPth, ToPth$)
ThwNotPth ToPth
ClrPthFil ToPth
Dim Ffn
For Each Ffn In FfnAy(FmPth)
    CpyFilzToPth Ffn, ToPth
Next
End Sub

Sub CpyFilzUp(Ffn)
CpyFilzToPth Ffn, ParPth(Ffn)
End Sub

Sub CpyFilzToNxtzAy(FfnAy$())
Dim I
For Each I In Itr(FfnAy)
    CpyFilzToNxt I
Next
End Sub

Function CpyFilzToNxt$(Ffn)
Dim O$
O = NxtFfn(Ffn)
CpyFilzToFfn Ffn, O
CpyFilzToNxt = O
End Function

Sub CpyFilzIfDif(FfnAy_or_Ffn, ToPth$, Optional UseEq As Boolean)
Dim I
For Each I In ItrzStr(FfnAy_or_Ffn)
    CpyFilzIfDifzSng I, ToPth, UseEq
Next
End Sub

Sub CpyFilzToFfn(FmFfn, ToFfn$, Optional OvrWrt As Boolean)
Fso.GetFile(FmFfn).Copy ToFfn, OvrWrt
End Sub

Function CpyFilzToPth$(FfnAy_or_Ffn, ToPth$, Optional OvrWrt As Boolean)
Dim Ffn, P$, O$
P = PthEnsSfx(ToPth)
For Each Ffn In ItrzStr(FfnAy_or_Ffn)
    O = P & Fn(Ffn)
    CpyFilzToFfn Ffn, O, OvrWrt
Next
CpyFilzToPth = O
End Function

Sub CpyFilzIfDifzSng(Ffn, ToPth$, Optional UseEq As Boolean)
Dim B$, Msg$, IsDif As Boolean
Select Case True
Case HasFfn(Ffn)
    B = ToPth & Fn(Ffn)
    IsDif = IsDifFfn(B, Ffn, UseEq)
    Select Case True
    Case IsDif: Fso.CopyFile Ffn, B, True
                Msg = "Fil copied": GoSub Prt
    Case Else:  Msg = FmtQQ("No Copy, file is ?.", IIf(UseEq, "eq", "same")): GoSub Prt
    End Select
Case Else
    Thw CSub, "File not found", "FmFfn ToPth IsDif HasToFfn UseEq", Ffn, ToPth, IsDif, HasFfn(B), UseEq
End Select
Exit Sub
Prt:
    Inf CSub, Msg, "Fil FmPth ToPth IsDif HasToFfn Si Tim", Fn(Ffn), Pth(Ffn), ToPth, IsDif, HasFfn(B), FfnSz(Ffn), FfnTimStr(Ffn)
    Return
End Sub

Function IsNxtFfn(Ffn) As Boolean
Dim Las5$: Las5 = Right(Fn(Ffn), 5)
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

Function FfnzNxtFfn$(NxtFfn)
If IsNxtFfn(NxtFfn) Then
    Dim F$: F = Fn(NxtFfn)
    FfnzNxtFfn = Pth(NxtFfn) & Left(F, Len(F) - 5) & Ext(NxtFfn)
Else
    FfnzNxtFfn = NxtFfn
End If
End Function

Function NxtFfn$(Ffn)
If Not HasFfn(Ffn) Then NxtFfn = Ffn: Exit Function
Dim J%, O$
For J = 1 To 999
    O = FfnAddFnSfx(Ffn, "(" & Format(J, "000") & ")")
    If Not HasFfn(O) Then NxtFfn = O: Exit Function
Next
Stop
End Function


