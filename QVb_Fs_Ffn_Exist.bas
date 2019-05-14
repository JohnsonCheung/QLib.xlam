Attribute VB_Name = "QVb_Fs_Ffn_Exist"
Option Explicit
Private Const CMod$ = "MVb_Fs_Ffn_Exist."
Private Const Asm$ = "QVb"

Sub AsgFfnExistMisAset(OExistFfn As Aset, OMisFfn As Aset, Ffny$())
With FfnExist_FmFfny_ToExistNy_AndMisNy(Ffny)
    Set OExistFfn = AsetzAy(.Sy1)
    Set OMisFfn = AsetzAy(.Sy2)
End With
End Sub

Sub FfnExist_FmFfny_ToExist_AndMis(Ffny$(), OExist$(), OMis$())
Dim Ffn
Erase OExist
Erase OMis
For Each Ffn In Itr(Ffny)
    If HasFfn(Ffn) Then
        PushI OExist, Ffn
    Else
        PushI OMis, Ffn
    End If
Next
End Sub
Function FfnywExist(Ffny$()) As String()
Dim Ffn$, I
For Each I In Itr(Ffny)
    Ffn = I
    If HasFfn(Ffn) Then
        PushI FfnywExist, Ffn
    End If
Next
End Function


Function HasFfn(Ffn) As Boolean
HasFfn = Fso.FileExists(Ffn)
End Function


Function ExistFfnAset(Ffny$()) As Aset
Set ExistFfnAset = AsetzAy(ExistFfny(Ffny))
End Function

Function MisFfnAset(Ffny$()) As Aset
Set MisFfnAset = AsetzAy(MisFfny(Ffny))
End Function

Function ExistFfny(Ffny$()) As String()
Dim Ffn$, I
For Each I In Itr(Ffny)
    Ffn = I
    If HasFfn(Ffn) Then PushI ExistFfny, Ffn
Next
End Function
Function MisFfny(Ffny$()) As String()
Dim Ffn$, I
For Each I In Itr(Ffny)
    Ffn = I
    If Not HasFfn(Ffn) Then PushI MisFfny, Ffn
Next
End Function

Function IsFfn(Ffn) As Boolean
IsFfn = Fso.FileExists(Ffn)
End Function


Function ChkHasFfn(Ffn, Optional FileKind$ = "File") As String()
If Not HasFfn(Ffn) Then ChkHasFfn = MsgzMisKdFil(KdFil(Ffn, FileKind))
End Function

Sub ThwIf_FfnNotExist(Ffn, Fun$, Optional KdFil$)
If Not HasFfn(Ffn) Then Thw Fun, "File not found", "File-Pth File-Name File-Kind", Pth(Ffn), Fn(Ffn), KdFil
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
Sub EnsFfn(Ffn)
If Not HasFfn(Ffn) Then WrtStr "", Ffn
End Sub


