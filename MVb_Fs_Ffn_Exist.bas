Attribute VB_Name = "MVb_Fs_Ffn_Exist"
Option Explicit

Sub AsgFfnExistMisAset(OExistFfn As Aset, OMisFfn As Aset, FfnSy$())
With FfnExistPair(FfnSy)
    Set OExistFfn = AsetzAy(.Sy1)
    Set OMisFfn = AsetzAy(.Sy2)
End With
End Sub

Function FfnExistPair(FfnSy$()) As SyPair
Dim Ffn$, I, Exist$(), NotE$()
For Each I In Itr(FfnSy)
    Ffn = I
    If HasFfn(Ffn) Then
        PushI Exist, Ffn
    Else
        PushI NotE, Ffn
    End If
Next
Set FfnExistPair = SyPair(Exist, NotE)
End Function
Function FfnAywExist(FfnSy$()) As String()
Dim Ffn$, I
For Each I In Itr(FfnSy)
    Ffn = I
    If HasFfn(Ffn) Then
        PushI FfnAywExist, Ffn
    End If
Next
End Function


Function HasFfn(Ffn$) As Boolean
HasFfn = Fso.FileExists(Ffn$)
End Function


Function ExistFfnAset(FfnSy$()) As Aset
Set ExistFfnAset = AsetzAy(ExistFfnAy(FfnSy))
End Function

Function MisFfnAset(FfnSy$()) As Aset
Set MisFfnAset = AsetzAy(MisFfnAy(FfnSy))
End Function

Function ExistFfnAy(FfnSy$()) As String()
Dim Ffn$, I
For Each I In Itr(FfnSy)
    Ffn = I
    If HasFfn(Ffn) Then PushS ExistFfnAy, Ffn
Next
End Function
Function MisFfnAy(FfnSy$()) As String()
Dim Ffn$, I
For Each I In Itr(FfnSy)
    Ffn = I
    If Not HasFfn(Ffn) Then PushI MisFfnAy, Ffn
Next
End Function

Function IsFfn(Ffn$) As Boolean
IsFfn = Fso.FileExists(Ffn)
End Function


Function ChkHasFfn(Ffn$, Optional FileKind$ = "File") As String()
If Not HasFfn(Ffn) Then ChkHasFfn = MsgzMisFfn(Ffn$, FileKind)
End Function

Sub ThwIfFfnNotExist(Ffn$, Fun$, Optional FilKd$)
If Not HasFfn(Ffn) Then Thw Fun, "File not found", "File-Pth File-Name File-Kind", Pth(Ffn$), Fn(Ffn$), FilKd
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
Sub EnsFfn(Ffn$)
If Not HasFfn(Ffn) Then WrtStr "", Ffn
End Sub


