Attribute VB_Name = "MxRes"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxRes."

Function ResStr$(ResFn$, Optional ResPseg$)
ResStr = JnCrLf(ResLy(ResFn, ResPseg))
End Function

Sub WrtRes(S$, ResFn$, Optional ResPseg$, Optional OvrWrt As Boolean)
Dim Ft$: Ft = ResFfn(ResFn, ResPseg)
WrtStr S, Ft, OvrWrt
End Sub

Function ResLy(ResFn$, Optional ResPseg$) As String()
ResLy = LyzFt(ResFfn(ResFn, ResPseg))
End Function

Function ResHom$()
ResHom = AddFdrEns(CPjPth, CPjfn & ".res")
End Function

Function ResPth$(ResPseg$)
ResPth = AddPsegEns(ResHom, ResPseg)
End Function

Function ResFfn$(ResFn$, Optional ResPseg$)
'Ret : :Ft #Resource-Ffn#
ResFfn = ResPth(ResPseg) & ResFn
End Function
