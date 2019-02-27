Attribute VB_Name = "MVb_Fs_Pth_Mbr"
Option Explicit

Function PthDir$(Pth)
PthDir = Dir(PthEnsSfx(Pth) & "*.*", vbDirectory)
End Function

Function FdrAyz(Pth, Optional Spec$ = "*.*", Optional Atr As VbFileAttribute) As String()
If Not HasPth(Pth) Then Exit Function
Dim P$: P = PthEnsSfx(Pth)
Dim M$, X&, Atr1&
X = Atr Or vbDirectory
M = Dir(P & Spec, vbDirectory)
While M <> ""
    If InStr(M, "?") > 0 Then
        Info CSub, "Unicode entry is skipped", "UniCode-Entry Pth Spec Atr", M, Pth, Spec, Atr
        GoTo Nxt
    End If
    If M = "." Then GoTo Nxt
    If M = ".." Then GoTo Nxt
    Atr1 = GetAttr(P & M)
    If (Atr1 And VbFileAttribute.vbDirectory) = 0 Then GoTo Nxt
    If Not HitFilAtr(GetAttr(P & M), Atr) Then GoTo Nxt
    PushI FdrAyz, M    '<====
Nxt:
    M = Dir
Wend
End Function


Function EntAy(Pth) As String()
'Function EntAy(A$, Optional FilSpec$ = "*.*", Optional Atr As FileAttribute) As String()
Dim A$: A$ = PthDir(Pth)
While A <> ""
    If A = "." Then GoTo X
    If A = ".." Then GoTo X
    PushI EntAy, A
X:
    A = Dir
Wend
End Function

Function FdrAy(Pth) As String()
Dim P$: P = PthEnsSfx(Pth)
Dim A$: A = PthDir(P)
While A <> ""
    If A = "." Then GoTo X
    If A = ".." Then GoTo X
    If IsPth(P & A) Then PushI FdrAy, A
X:
    A = Dir
Wend
End Function

Function FfnItr(Pth)
Asg Itr(FfnAy(Pth)), FfnItr
End Function

'Function FnAy(Pth) As String()
'Dim A$: A = Dir(Pth)
'While A <> ""
'    If HasSubStr(A, "?") Then
'        Debug.Print FmtQQ("File name has ?, skipped.  Pth[?] Fn[?]", Pth, A)
'    Else
'        PushI FnAy, A
'    End If
'    A = Dir
'Wend
'End Function
'
'Function FfnAy(Pth) As String()
'FfnAy = AyAddPfx(FnAy(Pth), PthEnsSfx(Pth))
'End Function
'

Function SubPthAy(Pth) As String()
SubPthAy = AyAddPfxSfx(FdrAy(Pth), Pth, PthSep)
End Function

Function SubPthAyz(Pth, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
SubPthAyz = AyAddPfxSfx(FdrAyz(Pth, Spec, Atr), Pth, PthSep)
End Function


Sub AsgEnt(OFdrAy$(), OFnAy$(), Pth$)
Erase OFdrAy
Erase OFnAy
Dim A$, P$
P = PthEnsSfx(Pth)
A = Dir(Pth, vbDirectory)
While A <> ""
    If A = "." Then GoTo X
    If A = ".." Then GoTo X
    If IsPth(P & A) Then
        PushI OFdrAy, A
    Else
        PushI OFnAy, A
    End If
    A = Dir
X:
Wend
End Sub
Function FnnAy(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
Dim Fn
For Each Fn In FnAy(A, Spec)
    PushI FnnAy, RmvExt(Fn)
Next
End Function

Function FnAy(Pth, Optional Spec$ = "*.*") As String()
ThwNotHasPth Pth, CSub
Dim O$()
Dim M$
M = Dir(PthEnsSfx(Pth) & Spec)
While M <> ""
   PushI FnAy, M
   M = Dir
Wend
End Function

Function FxAy(Pth) As String()
Dim O$(), B$, P$
P = PthEnsSfx(P)
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

Function FfnAy(Pth, Optional Spec$ = "*.*") As String()
FfnAy = AyAddPfx(FnAy(Pth, Spec), Pth)
End Function

Private Sub Z_SubPthAy()
Dim Pth$
Pth = "C:\Users\user\AppData\Local\Temp\"
Ept = Sy()
GoSub Tst
Exit Sub
Tst:
    Act = SubPthAy(Pth)
    Brw Act
    Return
End Sub

Private Sub ZZ_FxAy()
Dim A$()
A = FxAy(CurDir)
DmpAy A
End Sub

