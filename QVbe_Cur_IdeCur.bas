Attribute VB_Name = "QVbe_Cur_IdeCur"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Cur_CdPne_Md_Mth."
Private Const Asm$ = "QIde"
Property Get CMdn()
CMdn = CCmp.Name
End Property
Function MthLno(Md As CodeModule, Lno&)
Dim O&
For O = Lno To 1 Step -1
    If IsMthLin(Md.Lines(O, 1)) Then MthLno = O: Exit Function
Next
End Function
Function NMthLin%(M As CodeModule, MthLno&)
Dim K$, J&, N&, E$
K = MthKd(M.Lines(MthLno, 1))
If K = "" Then Thw CSub, "Given MthLno is not a MthLin", "Md MthLno MthLin", Mdn(M), MthLno, M.Lines(MthLno, 1)
E = "End " & K
For J = MthLno To M.CountOfLines
    N = N + 1
    If M.Lines(J, 1) = E Then NMthLin = N: Exit Function
Next
ThwImpossible CSub
End Function
Property Get CLno&()
CLno = LnozM(CMd)
End Property
Property Get CMthn$()
Dim M As CodeModule: Set M = CMd
If IsNothing(M) Then Exit Property
CMthn = MthnzM(CMd)
End Property
Function WinyzV(V As Vbe) As vbIde.Window
End Function
Function WinyV() As vbIde.Window()
WinyV = WinyzV(CVbe)
End Function

Function CWin() As vbIde.Window
Dim A As CodePane
Set A = CPne
If IsNothing(A) Then Exit Function
Set CWin = A.Window
End Function

Property Get CPne() As vbIde.CodePane
Set CPne = CVbe.ActiveCodePane
End Property
Property Get CMthLno&()
CMthLno = MthLnozM(CMd, CLno)
End Property

Property Get CMthLin$()
CMthLin = MthLinzML(CMd, CLno)
End Property

Property Get CVbe() As Vbe
Set CVbe = Application.Vbe
End Property


