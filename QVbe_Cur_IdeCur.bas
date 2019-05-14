Attribute VB_Name = "QVbe_Cur_IdeCur"
Option Explicit
Private Const CMod$ = "MIde_Cur_CdPne_Md_Mth."
Private Const Asm$ = "QIde"
Property Get CMdn()
CMdn = CCmp.Name
End Property
Property Get CLno&()
CLno = LnozM(CMd)
End Property
Property Get CMthn()
CMthn = MthnzM(CMd)
End Property

Property Get CCdWiny() As VBIDE.Window()
CCdWiny = Winyw(CWiny, vbext_wt_CodeWindow)
End Property

Property Get CWin() As VBIDE.Window
Dim A As CodePane
Set A = CPne
If IsNothing(A) Then Exit Property
Set CWin = A.Window
End Property

Property Get CPne() As VBIDE.CodePane
Set CPne = CVbe.ActiveCodePane
End Property
Property Get CMthLin()
Dim A As CodeModule: Set A = CMd
Dim Lno: Lno = CurLno
Dim J&
For J = Lno To 1 Step -1
    If IsMthLin(A.Lines(J, 1)) Then
        CMthLin = ContLinzML(A, J)
        Exit Property
    End If
Next
End Property
Property Get CVbe() As Vbe
Set CVbe = Application.Vbe
End Property


