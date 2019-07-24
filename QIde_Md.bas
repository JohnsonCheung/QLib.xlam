Attribute VB_Name = "QIde_Md"
Option Compare Text
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Md."
':DiMdnqSrc$ = "It is from Pj. Key is Mdn and Val is MdLines"
':MdDn$ = "Full: Md-Dot-Nm.  It is Either Mdn or Pjn-Dot-Mdn."
':MdDn: :Pjn.Mdn-or-Pjn
Function IsCls(M As CodeModule) As Boolean
IsCls = M.Parent.Type = vbext_ct_ClassModule
End Function

Function IsMod(M As CodeModule) As Boolean
IsMod = M.Parent.Type = vbext_ct_StdModule
End Function

Function MdzDn(MdDn) As CodeModule
Set MdzDn = Md(MdDn)
End Function

Function Md(MdDn) As CodeModule
Const CSub$ = CMod & "Md"
Dim A1$(): A1 = Split(MdDn, ".")
Select Case Si(A1)
Case 1: Set Md = CPj.VBComponents(A1(0)).CodeModule
Case 2: Set Md = Pj(A1(0)).VBComponents(A1(1)).CodeModule
Case Else: Thw CSub, "[MdDn] should be XXX.XXX or XXX", MdDn
End Select
End Function

Function MdDn$(M As CodeModule)
MdDn = PjnzM(M) & "." & Mdn(M)
End Function

Function DiMdnqSrc(P As VBProject) As Dictionary
Dim C As VBComponent
Set DiMdnqSrc = New Dictionary
For Each C In P.VBComponents
    DiMdnqSrc.Add C.Name, SrcL(C.CodeModule)
Next
End Function

Function DiMdnqSrcP() As Dictionary
Set DiMdnqSrcP = DiMdnqSrc(CPj)
End Function

Function MdFn$(M As CodeModule)
MdFn = Mdn(M) & ExtzCmpTy(CmpTyzM(M))
End Function

Function Mdn(M As CodeModule)
Mdn = M.Parent.Name
End Function

Function MdnzM(M As CodeModule)
MdnzM = M.Parent.Name
End Function

Function MdTy(M As CodeModule) As vbext_ComponentType
MdTy = M.Parent.Type
End Function

Function ShtCmpTyzM$(M As CodeModule)
ShtCmpTyzM = ShtCmpTy(CmpTyzM(M))
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
SizMd = Len(SrcL(M))
End Function

Function SrcL$(M As CodeModule)
SrcL = JnCrLf(Src(M)) & vbCrLf
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

Property Get CMdDn$()
CMdDn = MdDn(CMd)
End Property

Sub ClsMd(M As CodeModule)
M.CodePane.Window.Close
End Sub

Sub CmprMd(M As CodeModule, B As CodeModule)
'BrwCmpgDicAB DiMthnqLineszM(A), DiMthnqLineszMd(B), MdDn(A), MdDn(B)
End Sub

Private Function Y_Md() As CodeModule
Set Y_Md = CVbe.VBProjects("StockShipRate").VBComponents("Schm").CodeModule
End Function

Private Sub Z_DroMds()
'BrwDrs DroMds(Md("IdeFeature_EnsZ_AsPrivate"))
End Sub

Private Sub Z_MthLnozMM()
Dim O$()
    Dim Lno, L&(), M, A As CodeModule, Ny$(), J%
    Set A = Md("Fct")
    Ny = MthNyzM(A)
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

Function CvMd(A) As CodeModule
Set CvMd = A
End Function

Function LibNyP() As String()
LibNyP = LibNyzP(CPj)
End Function

Function LibNyzP(P As VBProject) As String()
LibNyzP = AeBlnk(AwDistAsSy(AyBef(MdNyzP(P), "_")))
End Function
