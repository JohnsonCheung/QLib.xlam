Attribute VB_Name = "QVb_Fs_Pth"
Option Compare Text
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Fs_Pth."
Function AddFdr$(Pth, Fdr)
AddFdr = EnsPthSfx(Pth) & AddNB(Fdr, "\")
End Function
Function AddPsegEns$(Pth, Pseg)
AddPsegEns = EnsPthzAllSeg(AddFdr(Pth, Pseg))
End Function

Function AddFdrEns$(Pth, Fdr)
AddFdrEns = EnsPth(AddFdr(Pth, Fdr))
End Function

Function AddFdrApEns$(Pth, ParamArray FdrAp())
Dim Av(): Av = FdrAp
Dim O$: O = AddFdrAv(Pth, Av)
EnsPthzAllSeg O
AddFdrApEns = O
End Function

Private Function AddFdrAv$(Pth, FdrAv())
Dim O$: O = Pth
Dim I, Fdr$
For Each I In FdrAv
    Fdr = I
    O = AddFdr(O, Fdr)
Next
AddFdrAv = O
End Function

Function AddFdrAp$(Pth, ParamArray FdrAp())
Dim Av(): Av = FdrAp
AddFdrAp = AddFdrAv(Pth, Av)
End Function

Function MsyzFfnAlreadyLoaded(Ffn, FilKind$, LTimStr$) As String()
Dim Si&, Tim$, Ld$, Msg$
Si = SizFfn(Ffn)
Tim = TimStrzFfn(Ffn)
Msg = FmtQQ("[?] file of [time] and [size] is already loaded [at].", FilKind)
MsyzFfnAlreadyLoaded = LyzMsgNap(Msg, Ffn, Tim, Si, LTimStr)
End Function

Function IsPthOfEmp(Pth) As Boolean
ThwIf_PthNotExist Pth, CSub
If AnyFil(Pth) Then Exit Function
If HasSubFdr(Pth) Then Exit Function
IsPthOfEmp = True
End Function

Function AddPfxzPth$(Pth, Pfx)
With Brk2Rev(RmvSfx(Pth, PthSep), PthSep, NoTrim:=True)
    AddPfxzPth = .S1 & PthSep & Pfx & .S2 & PthSep
End With
End Function

Function HitFilAtr(A As VbFileAttribute, Wh As VbFileAttribute) As Boolean
HitFilAtr = True
End Function

Function FdrzFfn$(Ffn)
FdrzFfn = Fdr(Pth(Ffn))
End Function
Function Fdr$(Pth)
Fdr = AftRev(RmvPthSfx(Pth), PthSep)
End Function

Sub ThwIf_NotProperFdrNm(Fdr$)
Const CSub$ = CMod & "ThwNotFdr"
Const C$ = "\/:<>"
If HasChrList(Fdr, C) Then Thw CSub, "Fdr cannot has these char " & C, "Fdr Char", Fdr, C
End Sub

Function PthRmvFdr$(Pth)
PthRmvFdr = BefRev(RmvPthSfx(Pth), PthSep) & PthSep
End Function

Function FfnUp$(Ffn)
FfnUp = PthRmvFdr(Pth(Ffn))
End Function


Function ParPth$(Pth) ' Return the ParPth of given Pth
If Not HasSubStr(Pth, PthSep) Then Err.Raise 1, "ParPth", "No PthSep in Pth" & vbCrLf & Pth
ParPth = BefRevOrAll(RmvLasChr(EnsPthSfx(Pth)), PthSep) & PthSep
End Function

Function ParFdr$(Pth)
ParFdr = Fdr(ParPth(Pth))
End Function

Function PthzUpN$(Pth, UpN%)
Dim O$, J%
O = Pth
For J = 1 To UpN
    O = ParPth(O)
Next
PthzUpN = O
End Function

Function EnsPth$(Pth)
Dim P$: P = EnsPthSfx(Pth)
If Not Fso.FolderExists(Pth) Then MkDir RmvLasChr(P)
EnsPth = Pth
End Function

Function EnsPthzAllSeg$(Pth)
'Ret : @Pth and ens each seg.
Dim J%, O$, Ay$()
Ay = Split(RmvSfx(Pth, PthSep), PthSep)
O = Ay(0)
For J = 1 To UBound(Ay)
    O = O & PthSep & Ay(J)
    EnsPth O
Next
End Function

Function HasPth(Pth) As Boolean
HasPth = IsPthExist(Pth)
End Function

Function HasFdr(Pth, Fdr$) As Boolean
HasFdr = HasEle(FdrAy(Pth), Fdr)
End Function

Sub ThwIf_PthNotExist(Pth, Fun$)
If Not HasPth(Pth) Then Thw Fun, "Pth not exist", "Pth", Pth
End Sub

Function AnyFil(Pth) As Boolean
AnyFil = Dir(Pth) <> ""
End Function
Function IsPth(Pth) As Boolean
IsPth = IsPthExist(Pth)
End Function
Function IsPthExist(Pth) As Boolean
IsPthExist = Fso.FolderExists(Pth)
End Function

Function HasSubFdr(Pth) As Boolean
HasSubFdr = Fso.GetFolder(Pth).SubFolders.Count > 0
End Function

Sub ThwIf_PthNotExist1(Pth, Optional Fun$ = "ThwIf_PthNotExist1")
If Not HasPth(Pth) Then Thw Fun, "Path not exist", "Path", Pth
End Sub


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

Function SubPthy(Pth) As String()
SubPthy = AddPfxSzAy(FdrAy(Pth), EnsPthSfx(Pth), PthSep)
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
Ffny = AddPfxzAy(FnAy(Pth, Spec), EnsPthSfx(Pth))
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

Private Sub Z_Fxy()
Dim A$()
A = Fxy(CurDir)
DmpAy A
End Sub



Function HasPthSfx(Pth) As Boolean
HasPthSfx = LasChr(Pth) = PthSep
End Function
Function EnsPthSfx$(Pth)
If HasPthSfx(Pth) Then
    EnsPthSfx = Pth
Else
    EnsPthSfx = Pth & PthSep
End If
End Function

Function RmvPthSfx$(Pth)
RmvPthSfx = RmvSfx(Pth, PthSep)
End Function




Function HasSiblingFdr(Pth, Fdr$) As Boolean
HasSiblingFdr = HasFdr(ParPth(Pth), Fdr)
End Function

Function SiblingPth$(Pth, SiblingFdr$)
SiblingPth = AddFdrEns(ParPth(Pth), SiblingFdr)
End Function



'
