Attribute VB_Name = "MVb_Fs_Ffn_Op_Cpy"
Option Explicit
Sub CpyPthzClr(FmPth, ToPth)
ThwNotPth ToPth
ClrPthFil ToPth
Dim Ffn
For Each Ffn In FfnAy(FmPth)
    CpyFilzToPth Ffn, ToPth
Next
End Sub
Function CpyFilzToPth$(Ffn, ToPth, Optional OvrWrt As Boolean)
ThwNotPth ToPth
ThwNoFfn Ffn, CSub, "File-to-backup"
Fso.CopyFile Ffn, ToPth, OvrWrt
CpyFilzToPth = ToPth & Fn(Ffn)
End Function

Sub CpyFilUp(Ffn)
CpyFilzToPth Ffn, ParPth(Ffn)
End Sub

Function CpyFfnToNxt$(Ffn)
Dim O$
O = NxtFfn(Ffn)
CpyFfnToFfn Ffn, O
CpyFfnToNxt = O
End Function


Sub CpyFfnAyToPthIfDif(FfnAy$(), Pth$)
Dim I
For Each I In FfnAy
    CpyFfnToPthIfDif I, Pth
Next
End Sub

Sub CpyFfnToFfn(FmFfn, ToFfn$, Optional OvrWrt As Boolean)
If OvrWrt Then DltFfn ToFfn
Dim F As File
Set F = Fso.GetFile(FmFfn)
F.Copy ToFfn
End Sub

Sub CpyFfnToPth(FmFfn, ToPth$, Optional OvrWrt As Boolean)
Dim ToFfn$: ToFfn = PthEnsSfx(ToPth) & Fn(FmFfn)
CpyFfnToFfn FmFfn, ToFfn
End Sub

Sub CpyFfnToPthIfDif(A, Pth$)
Dim B$, Msg$, IsSam As Boolean
Select Case True
Case HasFfn(A)
    B = Pth & Fn(A)
    IsSam = IsSamFfn(B, A)
    Select Case True
    Case IsSam:   Msg = "No Copy, file is same.": GoSub Prt
    Case Else:
        Fso.CopyFile A, B, True
        Msg = "Fil copied": GoSub Prt
    End Select
Case Else
    Thw CSub, "File not found", "FmFfn ToPth IsSam HasToFfn", A, Pth, IsSam, HasFfn(B)
End Select
Exit Sub
Prt:
    Info CSub, Msg, "FmFfn ToPth IsSam HasToFfn", A, Pth, IsSam, HasFfn(B)
    Return
End Sub


Function NxtFfn$(Ffn)
If Not HasFfn(Ffn) Then NxtFfn = Ffn: Exit Function
Dim J%, O$
For J = 1 To 999
    O = FfnAddFnSfx(Ffn, "(" & Format(J, "000") & ")")
    If Not HasFfn(O) Then NxtFfn = O: Exit Function
Next
Stop
End Function


