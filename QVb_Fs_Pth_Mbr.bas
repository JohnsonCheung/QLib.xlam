Attribute VB_Name = "QVb_Fs_Pth_Mbr"
Option Explicit
Private Const CMod$ = "MVb_Fs_Pth_Mbr."
Private Const Asm$ = "QVb"

Function DirzPth$(Pth$)
DirzPth = Dir(EnsPthSfx(Pth) & "*.*", vbDirectory)
End Function

Function FdrSyzIsInst(Pth$) As String()
Dim I, Fdr$
For Each I In Itr(FdrSy(Pth))
    Fdr = I
    If IsInstNm(Fdr) Then PushI FdrSyzIsInst, Fdr
Next
End Function
Function FdrSy(Pth$, Optional Spec$ = "*.*", Optional Atr As VbFileAttribute) As String()
If Not HasPth(Pth) Then Exit Function
Dim P$: P = EnsPthSfx(Pth)
Dim M$, X&, Atr1&
X = Atr Or vbDirectory
M = Dir(P & Spec, vbDirectory)
While M <> ""
    If InStr(M, "?") > 0 Then
        Inf CSub, "Unicode entry is skipped", "UniCode-Entry Pth Spec Atr", M, Pth, Spec, Atr
        GoTo Nxt
    End If
    If M = "." Then GoTo Nxt
    If M = ".." Then GoTo Nxt
    Atr1 = GetAttr(P & M)
    If (Atr1 And VbFileAttribute.vbDirectory) = 0 Then GoTo Nxt
    If Not HitFilAtr(GetAttr(P & M), Atr) Then GoTo Nxt
    PushI FdrSy, M    '<====
Nxt:
    M = Dir
Wend
End Function

Function EntSy(Pth$) As String()
'Function EntSy(A$, Optional FilSpec$ = "*.*", Optional Atr As FileAttribute) As String()
Dim A$: A$ = DirzPth(Pth)
While A <> ""
    If A = "." Then GoTo X
    If A = ".." Then GoTo X
    PushI EntSy, A
X:
    A = Dir
Wend
End Function
Function IsInstNm(S$) As Boolean
If FstChr(S) <> "N" Then Exit Function      'FstChr = N
If Len(S) <> 16 Then Exit Function          'Len    =16
If Not IsYYYYMMDD(Mid(S, 2, 8)) Then Exit Function 'NYYYYMMDD_HHMMDD
If Mid(S, 10, 1) <> "_" Then Exit Function
If Not IsHHMMDD(Right(S, 6)) Then Exit Function
IsInstNm = True
End Function

Function FdrSy1(Pth$) As String()
Dim P$: P = EnsPthSfx(Pth)
Dim A$: A = DirzPth(P)
While A <> ""
    If A = "." Then GoTo X
    If A = ".." Then GoTo X
    If IsPth(P & A) Then PushI FdrSy1, A
X:
    A = Dir
Wend
End Function

Function FfnItr(Pth$)
Asg Itr(FfnSy(Pth)), FfnItr
End Function

'Function FnSy(Pth) As String()
'Dim A$: A = Dir(Pth)
'While A <> ""
'    If HasSubStr(A, "?") Then
'        Debug.Print FmtQQ("File name has ?, skipped.  Pth[?] Fn[?]", Pth, A)
'    Else
'        PushI FnSy, A
'    End If
'    A = Dir
'Wend
'End Function
'
'Function FfnSy(Pth) As String()
'FfnSy = AddPfxzSy(FnSy(Pth), EnsPthSfx(Pth))
'End Function
'

Function SubPthSy(Pth$) As String()
SubPthSy = AddPfxzSySfx(FdrSy(Pth), Pth, PthSep)
End Function

Sub AsgEnt(OFdrSy$(), OFnAy$(), Pth$)
Erase OFdrSy
Erase OFnAy
Dim A$, P$
P = EnsPthSfx(Pth)
A = Dir(Pth, vbDirectory)
While A <> ""
    If A = "." Then GoTo X
    If A = ".." Then GoTo X
    If IsPth(P & A) Then
        PushI OFdrSy, A
    Else
        PushI OFnAy, A
    End If
    A = Dir
X:
Wend
End Sub

Function FnnSy(Pth$, Optional Spec$ = "*.*") As String()
Dim I
For Each I In FnSy(Pth, Spec)
    PushI FnnSy, RmvExt(CStr(I))
Next
End Function

Function FnSyzFfnSy(FfnSy$()) As String()
Dim I, Ffn$
For Each I In Itr(FfnSy)
    Ffn = I
    PushI FnSyzFfnSy, Fn(Ffn)
Next
End Function

Function FnSy(Pth$, Optional Spec$ = "*.*") As String()
ThwIfPthNotExist1 Pth, CSub
Dim O$()
Dim M$
M = Dir(EnsPthSfx(Pth) & Spec)
While M <> ""
   PushI FnSy, M
   M = Dir
Wend
End Function

Function FxAy(Pth$) As String()
Dim O$(), B$, P$
P = EnsPthSfx(P)
B = Dir(Pth & "*.xls")
Dim J%
While B <> ""
    J = J + 1
    If J > 1000 Then Stop
    If Ext(B) = ".xls" Then
        PushI O, Pth & B
    End If
    B = Dir
Wend
FxAy = O
End Function

Function FfnSy(Pth$, Optional Spec$ = "*.*") As String()
FfnSy = AddPfxzSy(FnSy(Pth, Spec), Pth)
End Function

Private Sub Z_SubPthSy()
Dim Pth$
Pth = "C:\Users\user\AppData\Local\Temp\"
Ept = Sy()
GoSub Tst
Exit Sub
Tst:
    Act = SubPthSy(Pth)
    Brw Act
    Return
End Sub

Private Sub ZZ_FxAy()
Dim A$()
A = FxAy(CurDir)
DmpAy A
End Sub
