Attribute VB_Name = "QVbe_Cur_IdeCur"
Option Compare Text
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
Property Get CMthLin()
Dim A As CodeModule: Set A = CMd
Dim Lno: Lno = CLno
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


