Attribute VB_Name = "QIde_Gen_GenModCmd"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Cmd."
Private Const Asm$ = "QIde"
Sub Cmd_SavGendMod()
Const C$ = "WsGenMod"
Const LoNm$ = "T_GenMod"
Dim Ws As Worksheet, Lo As ListObject, FxaCell As Range, Fxa$, Pj As VBProject, MdnCell As Range
Dim Mdn$, Md As CodeModule
Dim NewSrc$(), OldL$, NewL$
              If ChkNoWsCd(C) Then Exit Sub
     Set Ws = WszCd(C)
              If ChkNoLo(Ws, LoNm) Then Exit Sub
     Set Lo = Ws.ListObjects(LoNm)
              If ChkLoCCExact(Lo, "SrcCd") Then Exit Sub
Set FxaCell = Ws.Range("A1")
        Fxa = FxaCell.Value
              If ChkNotHasFfn(Fxa) Then FxaCell.Activate: Exit Sub
              OpnFxa Fxa
     Set Pj = PjzFxa(Fxa)
Set MdnCell = Ws.Range("A2")
        Mdn = MdnCell.Value
              If ChkMdn(Pj, Mdn) Then MdnCell.Activate: Exit Sub
              ActWs Ws
     Set Md = MdzPN(Pj, Mdn)
     NewSrc = StrColzWsLC(Lo, "T_GenMod", "SrcCd")
              If Si(NewSrc) = 0 Then MsgBox "No source code is generated", vbCritical: Exit Sub
       OldL = SrcLines(Md)
       NewL = JnCrLf(NewSrc)
              If OldL <> NewL Then MsgBox "The source in WsSrc <> the soruce from Mod", vbCritical: Exit Sub
              PutAyAtV NewSrc, A1zLo(Lo)     '<== Save
              MsgBox "SrcCd saved:" & vbCrLf & WrdCntg(NewL), vbInformation

End Sub
Sub Cmd_LoadSrcCd()
Const C$ = "WsSrcCd"
Const LoNm$ = "T_SrcCd"
Dim Ws As Worksheet
Dim Lo As ListObject
Dim FxaCell As Range
Dim Fxa$
Dim Mdn$
Dim Pj As VBProject
Dim MdnCell As Range
Dim Md As CodeModule
Dim SrcCd$()
              If ChkNoWsCd(C) Then Exit Sub
     Set Ws = WszCd(C)
              If ChkNoLo(Ws, LoNm) Then Exit Sub
     Set Lo = Ws.ListObjects(LoNm)
              If ChkLoCCExact(Lo, "SrcCd") Then Exit Sub
Set FxaCell = Ws.Range("A1")
        Fxa = FxaCell.Value
              If ChkNotHasFfn(Fxa) Then FxaCell.Activate: Exit Sub
              OpnFxa Fxa
     Set Pj = PjzFxa(Fxa)
Set MdnCell = Ws.Range("A2")
        Mdn = MdnCell.Value
              If ChkMdn(Pj, Mdn) Then MdnCell.Activate: Exit Sub
              ActWs Ws
     Set Md = MdzPN(Pj, Mdn)
      SrcCd = Src(Md)
              DltLoRow Lo                   '<== Delete
              PutAyAtV SrcCd, A1zLo(Lo)     '<== Load
              MsgBox "SrcCd loaded:" & vbCrLf & WrdCntg(JnCrLf(SrcCd)), vbInformation
End Sub
Sub Cmd_Srt()
'EnsWbNmzLcPfx WsSrc, "T_Src", "Key", "Stp"
End Sub

Sub Cmd_Vdt()
Const C$ = "WsGenMod"
Const LoNm$ = "T_SrcCd"
Dim ISrc As Drs
Dim IDes As Drs
Dim SrcCd$()
'
Dim Ws As Worksheet
Dim Lo As ListObject
Dim FxaCell As Range
Dim Fxa$
Dim Mdn$
Dim Pj As VBProject
Dim MdnCell As Range
Dim Md As CodeModule
              If ChkNoWsCd(C) Then Exit Sub
     Set Ws = WszCd(C)
              If ChkNoLo(Ws, LoNm) Then Exit Sub
     Set Lo = Ws.ListObjects(LoNm)
              If ChkLoCCExact(Lo, "SrcCd") Then Exit Sub
Set FxaCell = Ws.Range("A1")
        Fxa = FxaCell.Value
              If ChkNotHasFfn(Fxa) Then FxaCell.Activate: Exit Sub
              OpnFxa Fxa
     Set Pj = PjzFxa(Fxa)
Set MdnCell = Ws.Range("A2")
        Mdn = MdnCell.Value
              If ChkMdn(Pj, Mdn) Then MdnCell.Activate: Exit Sub
              ActWs Ws
      SrcCd = CdB(ISrc, IDes, SrcCd)
              DltLoRow Lo                   '<== Delete
              PutAyAtV SrcCd, A1zLo(Lo)     '<== Load
              MsgBox "SrcCd generated:" & vbCrLf & WrdCntg(JnCrLf(SrcCd)), vbInformation
End Sub

Sub Cmd_GenMod(ISrc As Drs, IDes As Drs, InpSrc$(), OupLo As ListObject)
Dim Src$(): Src = CdB(ISrc, IDes, InpSrc)
If HasWs(CWb, "Err") Then MsgBox "There is error", vbCritical: WszWb(CWb, "Err").Activate: Exit Sub
PutCd Src, OupLo
End Sub

Sub Cmd_EnsFmLnk()
Dim Rg As Range
'Set Rg = LoCC(WsSrc.ListObjects("T_Src"), "Fm1", "Fm5")
'EnsHypLnkzFollowNm Rg, "Stp"
End Sub

