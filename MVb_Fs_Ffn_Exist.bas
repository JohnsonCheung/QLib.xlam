Attribute VB_Name = "MVb_Fs_Ffn_Exist"
Option Explicit

Sub AsgFfnExistMisAset(OExistFfn As Aset, OMisFfn As Aset, FfnAy$())
With FfnExistPair(FfnAy)
    Set OExistFfn = AsetzAy(.Sy1)
    Set OMisFfn = AsetzAy(.Sy2)
End With
End Sub

Function FfnExistPair(FfnAy) As SyPair
Dim Ffn, Exist$(), NotE$()
For Each Ffn In Itr(FfnAy)
    If HasFfn(Ffn) Then
        PushI Exist, Ffn
    Else
        PushI NotE, Ffn
    End If
Next
Set FfnExistPair = SyPair(Exist, NotE)
End Function
Function FfnAywExist(FfnAy) As String()
Dim Ffn
For Each Ffn In Itr(FfnAy)
    If HasFfn(Ffn) Then
        PushI FfnAywExist, Ffn
    End If
Next
End Function


Function HasFfn(Ffn) As Boolean
HasFfn = Fso.FileExists(Ffn)
End Function


Function ExistFfnAset(FfnAy$()) As Aset
Set ExistFfnAset = AsetzAy(ExistFfnAy(FfnAy))
End Function

Function MisFfnAset(FfnAy$()) As Aset
Set MisFfnAset = AsetzAy(MisFfnAy(FfnAy))
End Function

Function ExistFfnAy(FfnAy$()) As String()
Dim Ffn
For Each Ffn In Itr(FfnAy)
    If HasFfn(Ffn) Then PushI ExistFfnAy, Ffn
Next
End Function
Function MisFfnAy(FfnAy$()) As String()
Dim Ffn
For Each Ffn In Itr(FfnAy)
    If Not HasFfn(Ffn) Then PushI MisFfnAy, Ffn
Next
End Function

Function IsFfn(Ffn) As Boolean
IsFfn = Fso.FileExists(Ffn)
End Function


Function ChkHasFfn(Ffn, Optional FileKind$ = "File") As String()
If Not HasFfn(Ffn) Then ChkHasFfn = MsgzMisFfn(Ffn, FileKind)
End Function

Sub ThwNoPth(Pth, Fun$, Optional PthKd$ = "Path")
If Not HasPth(Pth) Then Thw Fun, FmtQQ("? not found"), "Path", Pth
End Sub

Sub ThwNoFfn(Ffn, Fun$, Optional FilKd$)
If Not HasFfn(Ffn) Then Thw Fun, "File not found", "File-Pth File-Name File-Kind", Pth(Ffn), Fn(Ffn), FilKd
End Sub


Function LyzGpPth(Ffn As Aset) As String()
Dim P$(), F$()
    Dim J%, Ay$()
    Ay = Ffn.Srt.Sy
    For J = 0 To UB(Ay)
        PushI P, Pth(Ay(J))
        PushI F, Fn(Ay(J))
    Next
Dim O$()
    Dim LasP$, FstTim As Boolean
    FstTim = True
    For J = 0 To UB(P)
        If P(J) <> LasP Then
            If FstTim Then FstTim = False Else PushI O, ""
            PushI O, "Path: " & P(J)
            PushI O, "File: " & F(J)
            LasP = P(J)
        Else
            PushI O, "      " & F(J)
        End If
    Next

LyzGpPth = O
End Function


Function EnsFfn$(A)
If Not HasFfn(A) Then WrtStr "", A
EnsFfn = A
End Function


