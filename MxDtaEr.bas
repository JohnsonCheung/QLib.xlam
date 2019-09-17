Attribute VB_Name = "MxDtaEr"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxDtaEr."

Function EoColDup(D As Drs, C$) As String()
Dim B As Drs: B = F_SubDrs_ByDupFF(D, C)
Dim Msg$: Msg = "Dup [" & C & "]"
EoColDup = EoDrsMsg(B, Msg)
End Function

