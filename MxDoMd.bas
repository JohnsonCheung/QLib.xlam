Attribute VB_Name = "MxDoMd"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxDoMd."
Const FFoMd$ = "MdTy CLibv CNsv CModv Pjn Mdn NLin NMth NPub NPrv NFrd IsCModEr"
Function FoMd() As String()
FoMd = SyzSS(FFoMd)
End Function

Function DoMdP() As Drs
DoMdP = DoMdzP(CPj)
End Function

Private Function DoMdzP(P As VBProject) As Drs
DoMdzP = Drs(FoMd, DyoMdzP(P))
End Function

Private Function DroMd(M As CodeModule) As Variant()
Dim D(): D = DroMdn(M)
Dim CMod$: CMod = CModv(M)
Dim Pjn$: Pjn = D(0)
Dim MdTy$: MdTy = D(1)
Dim Mdn$: Mdn = D(2)
Dim IsCModEr As Boolean: IsCModEr = CMod <> Mdn
With MdSts(M)
    Dim NMth%: NMth = .NPub + .NPrv + .NFrd
    DroMd = Array(MdTy, CLibv(M), CNsv(M), CMod, Pjn, Mdn, .NLin, NMth, .NPub, .NPrv, .NFrd, IsCModEr)
End With
End Function

Private Function DyoMdzP(P As VBProject) As Variant()
Dim O()
    Dim C As VBComponent: For Each C In P.VBComponents
        Dim M As CodeModule: Set M = C.CodeModule
        PushI O, DroMd(M)
    Next
DyoMdzP = O
End Function

