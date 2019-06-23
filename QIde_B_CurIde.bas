Attribute VB_Name = "QIde_B_CurIde"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Vbe_Cur."
Private Const Asm$ = "QIde"

Function HasBar(BarNm) As Boolean
HasBar = HasBarzV(CVbe, BarNm)
End Function

Function HasPjf(Pjf) As Boolean
HasPjf = HasPjfzV(CVbe, Pjf)
End Function

Function PjzPjfC(Pjf) As VBProject
Set PjzPjfC = PjzPjf(CVbe, Pjf)
End Function

Function MdDrszV(A As Vbe) As Drs
MdDrszV = Drs(MdTblFny, MdDyoV(A))
End Function
Function MdTblFny() As String()

End Function

Sub SavCurVbe()
SavVbe CVbe
End Sub
Property Get CMdn()
CMdn = CCmp.Name
End Property
Function MthLno(Md As CodeModule, Lno&)
Dim O&
For O = Lno To 1 Step -1
    If IsLinMth(Md.Lines(O, 1)) Then MthLno = O: Exit Function
Next
End Function
Property Get CLno&()
CLno = LnozM(CMd)
End Property
Property Get CMthn$()
Dim M As CodeModule: Set M = CMd
If IsNothing(M) Then Exit Property
CMthn = CMthnzM(CMd)
End Property
Function WinyzV(V As Vbe) As VBIDE.Window
End Function
Function WinyV() As VBIDE.Window()
WinyV = WinyzV(CVbe)
End Function

Function CWin() As VBIDE.Window
Dim A As CodePane: Set A = CPne
If IsNothing(A) Then Exit Function
Set CWin = A.Window
End Function

Property Get CPne() As VBIDE.CodePane
Set CPne = CVbe.ActiveCodePane
End Property

Property Get CMthLno&()
CMthLno = MthLno(CMd, CLno)
End Property

Property Get CMthLin$()
CMthLin = MthLinzML(CMd, CLno)
End Property

Property Get CVbe() As Vbe
Set CVbe = Application.Vbe
End Property

Property Get CPjf$()
CPjf = CPj.Filename
End Property
Property Get CPj() As VBProject
Set CPj = CVbe.ActiveVBProject
End Property

Function HasMd(P As VBProject, Mdn, Optional IsInf As Boolean) As Boolean
Dim C As VBComponent
For Each C In P.VBComponents
    If C.Name = Mdn Then HasMd = True: Exit Function
Next
If IsInf Then
    Debug.Print FmtQQ("Mdn[?] not exist", Mdn)
End If
End Function

Sub ThwIf_NotMod(M As CodeModule, Fun$)
If Not IsMod(M) Then Thw Fun, "Should be a Mod", "Mdn MdTy", Mdn(M), ShtCmpTy(CmpTyzM(M))
End Sub

Function HasMod(P As VBProject, Modn) As Boolean
If Not HasMd(P, Modn) Then Exit Function
ThwIf_NotMod MdzPN(P, Modn), CSub
End Function
Function PjnyzX(X As Excel.Application) As String()
PjnyzX = PjnyzV(X.Vbe)
End Function
Property Get PjnyX() As String()
PjnyX = PjnyzX(Xls)
End Property
Property Get Pjn$()
Pjn = CPj.Name
End Property
Sub BrwPjp()
BrwPth PjpP
End Sub
Property Get CQMdn$()
CQMdn = MdDNm(CMd)
End Property

Property Get CQMthn$()
On Error GoTo X
CQMthn = QMthn(CMd, CMthLin)
Exit Property
X: Debug.Print CSub
End Property



