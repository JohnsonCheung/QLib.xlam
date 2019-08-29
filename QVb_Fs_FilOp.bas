Attribute VB_Name = "QVb_Fs_FilOp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Fs_Ffn_Backup."
Private Const Asm$ = "QVb"
Sub BrwBkp()
BrwPth BkpzP(CPj)
End Sub

Function BkpzP$(P As VBProject)
BkpzP = BkPth(Pjf(P))
End Function
Function BkRoot$(Pth)
BkRoot = AddFdr(Pth, ".Backup")
End Function

Function BkHom$(Ffn)
BkHom = AddFdr(BkRoot(Pth(Ffn)), Fn(Ffn))
End Function

Function LasBkFfn$(Ffn)
Dim H$: H = BkHom(Ffn)
Dim F$(): F = FdrAyzIsInst(H)
Dim Fdr$: Fdr = MaxEle(F)
LasBkFfn = H & Fdr & "\" & Fn(Ffn)
End Function

Function BkFfnAy(Ffn) As String()
Dim H$: H = BkHom(Ffn)
Dim F$(): F = FdrAyzIsInst(H)
Dim Fn1$: Fn1 = Fn(Ffn)
Dim Fdr: For Each Fdr In Itr(F)
    Dim IFfn$: IFfn = H & Fdr & "\" & Fn1
    If HasFfn(IFfn) Then
        PushI BkFfnAy, IFfn
    End If
Next
End Function

Function BkPjfAy() As String()
BkPjfAy = BkFfnAy(CPjf)
End Function

Function LasBkPjf$()
':FunAdj-Bk: :FunAdj #Backup#
':FunAdj-Las: :FunAdj #Last#
LasBkPjf = LasBkFfn(CPjf)
End Function

Function BkPth$(Ffn)
':BkPth: :Pth #Backup-Path# ! Backup path of a Ffn
BkPth = AddFdr(BkHom(Ffn), TmpNm)
End Function

Sub BackupP(Optional Msg$ = "Backup")
':FunVerb-Backup: :FunVerb
':FunSfx-P: :FunSfx #CurPj#
BackupFfn Pjf(CPj), Msg
End Sub

Function BkFfn$(Ffn)
BkFfn = BkPth(Ffn) & Fn(Ffn)
End Function

Function BackupFfn$(Ffn, Optional Msg$ = "Backup")
Dim T$: T = BkFfn(Ffn)
Dim Fmsg$: Fmsg = BkPth(Ffn) & "Msg.txt"
EnsPthzAllSeg Pth(T)
CpyFfn Ffn, T
WrtStr Msg, Fmsg
BackupFfn = T
End Function

Sub RplFfn(Ffn, ByFfn$)
BackupFfn Ffn
If DltFfnDone(Ffn) Then
    Name Ffn As ByFfn
End If
End Sub
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



Sub DltFfnyAyIf(Ffny$())
Dim Ffn
For Each Ffn In Itr(Ffny)
    DltFfnIf CStr(Ffn)
Next
End Sub

Sub DltFfn(Ffn)
On Error GoTo X
Kill Ffn
Exit Sub
X:
Thw CSub, "Cannot kill", "Ffn Er", Ffn, Err.Description
End Sub

Sub DltFfnIf(Ffn)
If HasFfn(Ffn) Then DltFfn Ffn
End Sub

Function DltFfnIfPrompt(Ffn, Msg$) As Boolean 'Return true if error
If Not HasFfn(Ffn) Then Exit Function
On Error GoTo X
Kill Ffn
Exit Function
X:
MsgBox "File [" & Ffn & "] cannot be deleted, " & vbCrLf & Msg
DltFfnIfPrompt = True
End Function

Function DltFfnDone(Ffn) As Boolean
On Error GoTo X
Kill Ffn
DltFfnDone = True
Exit Function
X:
End Function


Sub MovFilUp(Pth)
Dim I, Tar$
Tar$ = ParPth(Pth)
For Each I In Itr(FnAy(Pth))
    MovFfn CStr(I), Tar
Next
End Sub


Sub MovFfn(Ffn, ToPth$)
Fso.MoveFile Ffn, ToPth
End Sub




'
