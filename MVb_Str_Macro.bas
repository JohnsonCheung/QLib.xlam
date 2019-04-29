Attribute VB_Name = "MVb_Str_Macro"
Option Explicit

Function NyzMacro(A, Optional ExlBkt As Boolean, Optional OpnBkt$ = vbOpnBigBkt) As String()
'MacroStr-A is a with ..[xx].., this sub is to return all xx
Dim Q1$, Q2$
    Q1 = OpnBkt
    Q2 = ClsBkt(OpnBkt)
If Not HasSubStr(A, Q1) Then Exit Function

Dim O$(), J%
    Dim Ay$(): Ay = Split(A, Q1)
    For J = 1 To UB(Ay)
        Push O, Bef(Ay(J), Q2)
    Next
If Not ExlBkt Then
    O = SyAddPfxSfx(O, Q1, Q2)
End If
NyzMacro = O
End Function

Function FmtMacro(MacroStr$, ParamArray Ap()) As String()
Dim Av(): Av = Ap
FmtMacro = FmtMacroAv(MacroStr, Av)
End Function

Function FmtMacroAv(MacroStr$, Av()) As String()
FmtMacroAv = LyzNyAv(NyzMacro(MacroStr), Av)
End Function

Function FmtMacroDic$(MacroStr$, Dic As Dictionary)
Dim O$: O = MacroStr
Dim I, N$, K$
For Each I In Itr(NyzMacro(MacroStr))
    N = I
    K = RmvFstLasChr(N)
    If Dic.Exists(K) Then
        O = Replace(O, N, Dic(K))
    End If
Next
FmtMacroDic = O
End Function
