Attribute VB_Name = "MxDoMd"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDoMd."
Public Const FFoMd$ = "MdTy CLibv CNsv CModv Pjn Mdn IsCModEr NLin NMth NPub NPrv NFrd Mthnn"
Function FoMd() As String()
FoMd = SyzSS(FFoMd)
End Function

Function DoMdP() As Drs
DoMdP = DoMdzP(CPj)
End Function

Private Function DoMdzP(P As VBProject) As Drs
Dim MdId As Drs: MdId = DoMdIdzP(P)
DoMdzP = AddCol_MdSts_Mthnn(MdId, P)
End Function

Private Function AddCol_MdSts_Mthnn(MdId As Drs, P As VBProject) As Drs
Dim Dy()
    Dim IxMdn%: IxMdn = IxzAy(MdId.Fny, "Mdn")
    Dim Dr: For Each Dr In Itr(MdId.Dy)
        Dim Mdn$: Mdn = Dr(IxMdn)
        Dim M As CodeModule: Set M = P.VBComponents(Mdn).CodeModule
        Dim S$(): S = Src(M)
        Dim L$(): L = MthLinAy(S)
        With MdStszL(L)
            Dim NMth%: NMth = .NPub + .NPrv + .NFrd
            PushI Dy, AddAy(Dr, Array(.NLin, NMth, .NPub, .NPrv, .NFrd, MthnnzL(L)))
        End With
    Next
AddCol_MdSts_Mthnn = AddColzFFDy(MdId, "NLin NMth NPub NPrv NFrd Mthnn", Dy)
End Function
