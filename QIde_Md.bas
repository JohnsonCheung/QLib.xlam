Attribute VB_Name = "QIde_Md"
Option Compare Text
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Md."
Public Const DoczMdDic$ = "It is from Pj. Key is Mdn and Val is MdLines"
Public Const DoczMdDNm$ = "Full: Md-Dot-Nm.  It is Either Mdn or Pjn-Dot-Mdn."
Function IsCls(M As CodeModule) As Boolean
IsCls = M.Parent.Type = vbext_ct_ClassModule
End Function

Function IsMod(M As CodeModule) As Boolean
IsMod = M.Parent.Type = vbext_ct_StdModule
End Function

Function MdzDNm(MdDNm) As CodeModule
Set MdzDNm = Md(MdDNm)
End Function
Function Md(MdDNm) As CodeModule
Const CSub$ = CMod & "Md"
Dim A1$(): A1 = Split(MdDNm, ".")
Select Case Si(A1)
Case 1: Set Md = CPj.VBComponents(A1(0)).CodeModule
Case 2: Set Md = Pj(A1(0)).VBComponents(A1(1)).CodeModule
Case Else: Thw CSub, "[MdDNm] should be XXX.XXX or XXX", MdDNm
End Select
End Function

Function CmpAyzM(A() As CodeModule) As VBComponent()
CmpAyzM = IntozOP(CmpAyzM, A, PrpPth("Parent"))
End Function

Function MdAywInTy(A() As CodeModule, InTyAy() As vbext_ComponentType) As CodeModule()
Dim I
For Each I In A
    If HasEle(InTyAy, CvMd(I).Parent.Type) Then PushObj MdAywInTy, I
Next
End Function

Function MdDNm$(M As CodeModule)
MdDNm = PjnzM(M) & "." & Mdn(M)
End Function

Function SMdDiczP(P As VBProject) As Dictionary
Set SMdDiczP = SrtDic(MdDic(P))
End Function
Function MdDic(P As VBProject) As Dictionary
Dim C As VBComponent
Set MdDic = New Dictionary
For Each C In P.VBComponents
    MdDic.Add C.Name, SrcLines(C.CodeModule)
Next
End Function

Function MdDicP() As Dictionary
Set MdDicP = MdDic(CPj)
End Function

Function MdFn$(M As CodeModule)
MdFn = Mdn(M) & ExtzCmpTy(CmpTyzM(M))
End Function

Function Mdn(M As CodeModule)
Mdn = M.Parent.Name
End Function

Function QMdnzM$(M As CodeModule)
QMdnzM = PjnzM(M) & "." & Mdn(M)
End Function

Function MdTy(M As CodeModule) As vbext_ComponentType
MdTy = M.Parent.Type
End Function

Function ShtCmpTyzM$(M As CodeModule)
ShtCmpTyzM = ShtCmpTy(CmpTyzM(M))
End Function

Function NUsrTyMd%(M As CodeModule)
NUsrTyMd = NUsrTySrc(DclLyzM(M))
End Function

Function PjnzC(A As VBComponent)
PjnzC = A.Collection.Parent.Name
End Function

Function PjnzM(M As CodeModule)
PjnzM = PjnzC(M.Parent)
End Function

Function PjzM(M As CodeModule) As VBProject
Set PjzM = M.Parent.Collection.Parent
End Function

Function SizMd&(M As CodeModule)
SizMd = Len(SrcLines(M))
End Function

Function SrcLines$(M As CodeModule)
SrcLines = JnCrLf(Src(M)) & vbCrLf
End Function

Function RmvMthInSrc(Src$(), MthnSet As Aset) As String()
Dim D As Dictionary: 'Set D = DiceKeySet(MthnDic(Src), MthnSet): 'Brw D: Stop
RmvMthInSrc = LyzLinesDicItems(D)
End Function

Property Get CMd() As CodeModule
Dim P As CodePane: Set P = CPne
If IsNothing(P) Then Exit Property
Set CMd = P.CodeModule
End Property

Property Get CMdDNm$()
CMdDNm = QMdnzM(CMd)
End Property

Sub ClsMd(M As CodeModule)
M.CodePane.Window.Close
End Sub

Sub CmprMd(M As CodeModule, B As CodeModule)
'BrwCmpgDicAB MthDiczM(A), MthDiczMd(B), QMdnzM(A), QMdnzM(B)
End Sub

Sub DltLin(M As CodeModule, Lno)
M.DeleteLines Lno, 1
End Sub

Private Function Y_Md() As CodeModule
Set Y_Md = CVbe.VBProjects("StockShipRate").VBComponents("Schm").CodeModule
End Function

Private Sub ZZ_MdDrs()
'BrwDrs MdDrs(Md("IdeFeature_EnsZZ_AsPrivate"))
End Sub

Private Sub ZZ_MthLnozMM()
Dim O$()
    Dim Lno, L&(), M, A As CodeModule, Ny$(), J%
    Set A = Md("Fct")
    Ny = MthnyzM(A)
    For Each M In Ny
        DoEvents
        J = J + 1
        Push L, MthLnozMM(A, CStr(M))
        If J Mod 150 = 0 Then
            Debug.Print J, Si(Ny), "Z_MthLnozMM"
        End If
    Next

    For Each Lno In L
        Push O, Lno & " " & A.Lines(Lno, 1)
    Next
BrwAy O
End Sub

Private Sub Z_CMd()
Ass CMd.Parent.Name = "Cur_d"
End Sub
