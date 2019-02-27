Attribute VB_Name = "MVb_Fs_Ffn_MisEr"
Option Explicit
Const CMod$ = ""
Sub ThwMisFfnAy(FfnAy$(), Fun$, Optional FilKind$ = "File")
Const CSub$ = CMod & "ThwMisFfnAy"
ThwErMsg ChkMisFfnAy(FfnAy, FilKind), CSub, "Some given files not found", "Given files", FfnAy
End Sub

Function MsgzMisFfn(Ffn, Optional FilKind$ = "File") As String()
Dim F$
F = FmtQQ("[?] not found", FilKind)
MsgzMisFfn = LyzMsgNap(F, "File Path", Fn(Ffn), Pth(Ffn))
End Function

Function MsgzMisFfnAset(MisFfn As Aset, Optional FilKind$ = "file") As String()
If MisFfn.IsEmp Then Exit Function
Dim F$, S$
If MisFfn.Cnt = 1 Then
    S = FmtQQ("one ?", FilKind)
Else
    S = FmtQQ("? ?s", MisFfn.Cnt, FilKind)
End If
PushI MsgzMisFfnAset, FmtQQ("Following ? not found", S)
PushIAy MsgzMisFfnAset, AyTab(LyzGpPth(MisFfn))
End Function

Function MsgzMisFfnAy(FfnAy$(), Optional FilKind$ = "File") As String()
MsgzMisFfnAy = MsgzMisFfnAset(AsetzAy(FfnAy), FilKind)
End Function

Function ChkMisFfn(Ffn$, Optional FilKind$ = "File") As String()
If HasFfn(Ffn) Then Exit Function
ChkMisFfn = MsgzMisFfn(Ffn, FilKind)
End Function

Function ChkMisFfnAy(FfnAy$(), Optional FilKind$ = "File") As String()
Dim I, O$()
For Each I In FfnAy
    If Not HasFfn(I) Then
        PushI O, I
    End If
Next
ChkMisFfnAy = MsgzMisFfnAy(O, FilKind)
End Function
