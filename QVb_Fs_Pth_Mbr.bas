Attribute VB_Name = "QVb_Fs_Pth_Mbr"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Fs_Pth_Mbr."
Private Const Asm$ = "QVb"

Function DirzPSA$(Pth, Optional Spec$ = "*.*", Optional A As VbFileAttribute = VbFileAttribute.vbDirectory)
DirzPSA = Dir(EnsPthSfx(Pth) & Spec, A)
End Function

Function FdrAyzIsInst(Pth) As String()
Dim I, Fdr$
For Each I In Itr(FdrAy(Pth))
    Fdr = I
    If IsInstNm(Fdr) Then PushI FdrAyzIsInst, Fdr
Next
End Function

Function FdrAy(Pth, Optional Spec$ = "*.*") As String()
If Not HasPth(Pth) Then Exit Function
Dim P$: P = EnsPthSfx(Pth)
Dim F, X&, Atr1&
For Each F In Itr(EntAy(P, Spec))
    Atr1 = GetAttr(P & F)
    If (Atr1 And VbFileAttribute.vbDirectory) <> 0 Then
        PushI FdrAy, F    '<====
    End If
Next
End Function

Function EntAy(Pth, Optional Spec$ = "*.*", Optional Atr As FileAttribute = vbDirectory) As String()
Dim A$: A$ = DirzPSA(EnsPthSfx(Pth), Spec, Atr)
While A <> ""
    If A = "." Then GoTo X
    If A = ".." Then GoTo X
    If InStr(A, "?") > 0 Then
        Inf CSub, "Unicode entry is skipped", "UniCode-Entry Pth Spec", A, Pth, Spec
        GoTo X
    End If
    PushI EntAy, A
X:
    A = Dir
Wend
End Function
Function IsInstNm(Nm) As Boolean
If FstChr(Nm) <> "N" Then Exit Function      'FstChr = N
If Len(Nm) <> 16 Then Exit Function          'Len    =16
If Not IsYYYYMMDD(Mid(Nm, 2, 8)) Then Exit Function 'NYYYYMMDD_HHMMDD
If Mid(Nm, 10, 1) <> "_" Then Exit Function
If Not IsHHMMDD(Right(Nm, 6)) Then Exit Function
IsInstNm = True
End Function

Function FdrAy1(Pth) As String()
Dim P$: P = EnsPthSfx(Pth)
Dim A$: A = DirzPSA(P)
While A <> ""
    If A = "." Then GoTo X
    If A = ".." Then GoTo X
    If IsPth(P & A) Then PushI FdrAy1, A
X:
    A = Dir
Wend
End Function

Function FfnItr(Pth)
Asg Itr(Ffny(Pth)), FfnItr
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
'Function Ffny(Pth) As String()
'Ffny = AddPfxzAy(FnAy(Pth), EnsPthSfx(Pth))
'End Function
'

Function SubPthy(Pth) As String()
SubPthy = AddPfxSfxzAy(FdrAy(Pth), Pth, PthSep)
End Function

Sub AsgEnt(OFdrAy$(), OFnAy$(), Pth)
Erase OFdrAy
Erase OFnAy
Dim A$, P$
P = EnsPthSfx(Pth)
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

Function Fnny(Pth, Optional Spec$ = "*.*") As String()
Dim I
For Each I In FnAy(Pth, Spec)
    PushI Fnny, RmvExt(CStr(I))
Next
End Function

Function FnAyzFfny(Ffny$()) As String()
Dim I, Ffn$
For Each I In Itr(Ffny)
    Ffn = I
    PushI FnAyzFfny, Fn(Ffn)
Next
End Function

Function FnAy(Pth, Optional Spec$ = "*.*") As String()
ThwIf_PthNotExist1 Pth, CSub
Dim O$()
Dim M$
M = Dir(EnsPthSfx(Pth) & Spec)
While M <> ""
   PushI FnAy, M
   M = Dir
Wend
End Function

Function Fxy(Pth) As String()
Fxy = Ffny(Pth, "*.xls*")
End Function

Function Ffny(Pth, Optional Spec$ = "*.*") As String()
Ffny = AddPfxzAy(FnAy(Pth, Spec), Pth)
End Function

Private Sub Z_SubPthy()
Dim Pth
Pth = "C:\Users\user\AppData\Local\Temp\"
Ept = Sy()
GoSub Tst
Exit Sub
Tst:
    Act = SubPthy(Pth)
    Brw Act
    Return
End Sub

Private Sub ZZ_Fxy()
Dim A$()
A = Fxy(CurDir)
DmpAy A
End Sub
