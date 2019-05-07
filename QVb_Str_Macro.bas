Attribute VB_Name = "QVb_Str_Macro"
Option Explicit
Private Const CMod$ = "MVb_Str_Macro."
Private Const Asm$ = "QVb"
Type NNAv
    NN As String
    Av() As Variant
End Type
Function NyzMacro(Macro$, Optional OpnBkt$ = vbOpnBigBkt) As String()
'Macro is a str with ..[xx].., this sub is to return all xx
Dim Q1$:   Q1 = OpnBkt
Dim Q2$:   Q2 = ClsBkt(OpnBkt)
Dim Sy$(): Sy = Split(Macro, Q1)
NyzMacro = AywDist(RmvBlankLin(BefSy(Sy, Q2)))
End Function

Function FmtMacro$(Macro$, ParamArray Ap())
Dim Av(): Av = Ap
FmtMacro = FmtMacroAv(Macro, Av)
End Function

Function FmtMacroAv$(Macro$, Av())
Dim O$: O = Macro
Dim N, J%
For Each N In NyzMacro(Macro)
    O = Replace(O, "{" & N & "}", Av(J))
    J = J + 1
Next
FmtMacroAv = O
End Function
Function FmtMacrozRs$(Macro$, Rs As DAO.Recordset)
FmtMacroRs = FmtMacroDic(Macro, DiczRs(Rs))
End Function
Function DiczRs(A As DAO.Recordset) As Dictionary
Set DiczRs = New Dictionary
Dim F As DAO.Field
For Each F In A.Fields
    DiczRs.Add F.Name, F.Value
Next
End Function
Function FmtMacroDic$(Macro$, Dic As Dictionary)
With NNAv(Dic)
FmtMacrozDic = FmtMacroAv(Macro, .NN, .Av())
End With
End Function
Function NNAv(NN$, Av()) As NNAv
Dim N$(): N = Ny(NN)
ThwIfNotNy N, CSub
If Si(N) <> Si(Av) Then Thw CSub, "NN-Si <> Av-Si", "NN-Si Av-Si NN", Si(N), Si(Av), NN
NNAv.NN = NN
NNAv.Av = Av
End Function
Function NNAvzDic(A As Dictionary) As NNAv
NNAv.NN = JnSpc(NyzItr(A.Keys))
NNAv.Av = AvzItr(A.Items)
End With
End Function
Sub ThwIfNotNy(Ny$(), Fun$)
Dim N
For Each N In Itr(Ny)
    If Not IsNm(CStr(N)) Then Thw Fun, "Ele of Sy is not nm", "Not-nm-Ele Sy", N, Sy
Next
End Sub
