Attribute VB_Name = "MVb_Fs_Ffn_MisEr"
Option Explicit
Const CMod$ = ""
Sub ThwMisFfnAy(FfnSy$(), Fun$, Optional FilKind$ = "File")
Const CSub$ = CMod & "ThwMisFfnAy"
'ThwIfEr ChkMisFfnAy(FfnSy, FilKind), CSub, "Some given files not found", "Given files", FfnSy
End Sub

Function MsgzMisFfn(Ffn$, Optional FilKind$ = "File") As String()
Dim F$
F = FmtQQ("[?] not found", FilKind)
MsgzMisFfn = LyzMsgNap(F, "File Path", Fn(Ffn$), Pth(Ffn$))
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

Function MsgzMisFfnAy(FfnSy$(), Optional FilKind$ = "File") As String()
MsgzMisFfnAy = MsgzMisFfnAset(AsetzAy(FfnSy), FilKind)
End Function

Function ChkMisFfn(Ffn$, Optional FilKind$ = "File") As String()
If HasFfn(Ffn$) Then Exit Function
ChkMisFfn = MsgzMisFfn(Ffn$, FilKind)
End Function

Function ChkMisFfnAy(FfnSy$(), Optional FilKind$ = "File") As String()
Dim I, O$(), Ffn$
For Each I In FfnSy
    Ffn = I
    If Not HasFfn(Ffn) Then
        PushI O, I
    End If
Next
ChkMisFfnAy = MsgzMisFfnAy(O, FilKind)
End Function
