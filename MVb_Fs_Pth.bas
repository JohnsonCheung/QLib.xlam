Attribute VB_Name = "MVb_Fs_Pth"
Option Explicit
Const CMod$ = "MVb_Fs_Pth."
Private Function AddFdrzOne$(Pth, Fdr)
AddFdrzOne = PthEnsSfx(Pth) & Fdr & "\"
End Function

Function AddFdrEns$(Pth, ParamArray FdrAp())
Dim Av(): Av = FdrAp
AddFdrEns = PthEnsAll(AddFdrAv(Pth, Av))
End Function

Private Function AddFdrAv$(Pth, FdrAv())
Dim O$: O = Pth
Dim I
For Each I In FdrAv
    O = AddFdrzOne(O, I)
Next
AddFdrAv = O
End Function

Function AddFdr$(Pth, ParamArray FdrAp())
Dim Av(): Av = FdrAp
Dim O$: O = Pth
Dim I
For Each I In Itr(Av)
    O = AddFdrzOne(O, I)
Next
AddFdr = O
End Function

Function MsgzFfnAlreadyLoaded(Ffn$, FilKind$, LTimStr$) As String()
Dim Sz&, Tim$, Ld$, Msg$
Sz = FfnSz(Ffn)
Tim = FfnTimStr(Ffn)
Msg = FmtQQ("[?] file of [time] and [size] is already loaded [at].", FilKind)
MsgzFfnAlreadyLoaded = LyzMsgNap(Msg, Ffn, Tim, Sz, LTimStr)
End Function

Function IsEmpPth(Pth) As Boolean
ThwNotPth Pth
If HasFilPth(Pth) Then Exit Function
If HasSubFdr(Pth) Then Exit Function
IsEmpPth = True
End Function

Function PthAddPfx$(Pth, Pfx)
With Brk2Rev(RmvSfx(Pth, PthSep), PthSep, NoTrim:=True)
    PthAddPfx = .S1 & PthSep & Pfx & .S2 & PthSep
End With
End Function

Function HitFilAtr(A As VbFileAttribute, Wh As VbFileAttribute) As Boolean
HitFilAtr = True
End Function

Function Fdr$(Pth)
Fdr = TakAftRev(RmvLasChr(PthEnsSfx(Pth)), PthSep)
End Function

Sub ThwNotFdr(A)
Const CSub$ = CMod & "ThwNotFdr"
Const C$ = "\/:<>"
If HasChrList(A, C) Then Thw CSub, "Fdr cannot has these char " & C, "Fdr Char", A, C
End Sub

