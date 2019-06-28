Attribute VB_Name = "QVb_Str_Macro"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Str_Macro."
Private Const Asm$ = "QVb"
Type NNAv
    NN As String
    Av() As Variant
End Type

Function NyzMacro(Macro, Optional OpnBkt$ = vbOpnBigBkt, Optional InclBkt As Boolean) As String()
'Macro is a str with ..[xx].., it is to return all xx or [xx]
Dim Q1$:   Q1 = OpnBkt
Dim Q2$:   Q2 = ClsBkt(OpnBkt)
Dim Sy$(): Sy = Split(Macro, Q1)
Dim O$():   O = AwDist(RmvBlnkLin(BefzSy(Sy, Q2)))
If InclBkt Then O = AddPfxSzAy(O, Q1, Q2)
NyzMacro = O
End Function

Function FmtMacro(Macro, ParamArray Ap())
Dim Av(): Av = Ap
FmtMacro = FmtMacrozAv(Macro, Av)
End Function

Function FmtMacrozAv$(Macro, Av())
Dim O$: O = Macro
Dim N, J%
For Each N In NyzMacro(Macro)
    O = Replace(O, "{" & N & "}", Av(J))
    J = J + 1
Next
FmtMacrozAv = O
End Function
Function FmtMacrozRs$(Macro, Rs As Dao.Recordset)
FmtMacrozRs = FmtMacrozDic(Macro, DiczRs(Rs))
End Function
Function DiczRs(A As Dao.Recordset) As Dictionary
Set DiczRs = New Dictionary
Dim F As Dao.Field
For Each F In A.Fields
    DiczRs.Add F.Name, F.Value
Next
End Function
Function FmtMacrozDic$(Macro, Dic As Dictionary)
Dim Ny$(): Ny = NyzMacro(Macro)
Dim Vy(): Vy = VyzDicKy(Dic, Ny)
FmtMacrozDic = FmtMacrozAv(Macro, Vy)
End Function
Function NNAv(NN$, Av()) As NNAv
Dim N$(): N = Ny(NN)
ThwIf_NotNy N, CSub
If Si(N) <> Si(Av) Then Thw CSub, "NN-Si <> Av-Si", "NN-Si Av-Si NN", Si(N), Si(Av), NN
NNAv.NN = NN
NNAv.Av = Av
End Function

Function NNAvzDic(A As Dictionary) As NNAv
NNAvzDic.NN = JnSpc(NyzItr(A.Keys))
NNAvzDic.Av = AvzItr(A.Items)
End Function
Sub ThwIf_NotNy(Ny$(), Fun$)
Dim N
For Each N In Itr(Ny)
    If Not IsNm(N) Then Thw Fun, "Ele of Sy is not nm", "Not-nm-Ele Sy", N, Sy
Next
End Sub
