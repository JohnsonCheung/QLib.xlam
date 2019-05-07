Attribute VB_Name = "QVb_Fs_Pth"
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Fs_Pth."
Function AddFdr$(Pth$, Fdr$)
AddFdr = EnsPthSfx(Pth) & ApdIf(Fdr, "\")
End Function

Function AddFdrEns$(Pth$, Fdr$)
Dim O$: O = AddFdr(Pth, Fdr)
EnsPth O
AddFdrEns = O
End Function

Function AddFdrApEns$(Pth$, ParamArray FdrAp())
Dim Av(): Av = FdrAp
Dim O$: O = AddFdrAv(Pth, Av)
EnsPthzAllSeg O
AddFdrApEns = O
End Function

Private Function AddFdrAv$(Pth$, FdrAv())
Dim O$: O = Pth
Dim I, Fdr$
For Each I In FdrAv
    Fdr = I
    O = AddFdr(O, Fdr)
Next
AddFdrAv = O
End Function

Function AddFdrAp$(Pth$, ParamArray FdrAp())
Dim Av(): Av = FdrAp
AddFdrAp = AddFdrAv(Pth, Av)
End Function

Function MsyzFfnAlreadyLoaded(Ffn$, FilKind$, LTimStr$) As String()
Dim Si&, Tim$, Ld$, Msg$
Si = SizFfn(Ffn$)
Tim = DteTimStrzFfn(Ffn$)
Msg = FmtQQ("[?] file of [time] and [size] is already loaded [at].", FilKind)
MsyzFfnAlreadyLoaded = LyzMsgNap(Msg, Ffn, Tim, Si, LTimStr)
End Function

Function IsEmpPth(Pth$) As Boolean
ThwIfPthNotExist Pth
If AnyFil(Pth) Then Exit Function
If HasSubFdr(Pth) Then Exit Function
IsEmpPth = True
End Function

Function PthAddPfx$(Pth$, Pfx$)
With Brk2Rev(RmvSfx(Pth, PthSep), PthSep, NoTrim:=True)
    PthAddPfx = .S1 & PthSep & Pfx & .S2 & PthSep
End With
End Function

Function HitFilAtr(A As VbFileAttribute, Wh As VbFileAttribute) As Boolean
HitFilAtr = True
End Function

Function FdrzFfn$(Ffn$)
FdrzFfn = Fdr(Pth(Ffn$))
End Function

Function Fdr$(Pth$)
Fdr = AftRev(RmvPthSfx(Pth), PthSep)
End Function

Sub ThwIfNotProperFdrNm(Fdr$)
Const CSub$ = CMod & "ThwNotFdr"
Const C$ = "\/:<>"
If HasChrList(Fdr, C) Then Thw CSub, "Fdr cannot has these char " & C, "Fdr Char", Fdr, C
End Sub
