Attribute VB_Name = "QIde_Md"
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Md."
Public Const DoczMdDic$ = "It is from Pj. Key is Mdn and Val is MdLines"
Public Const DoczMdDNm$ = "Full: Md-Dot-Nm.  It is Either Mdn or Pjn-Dot-Mdn."
Function IsCls(A As CodeModule) As Boolean
IsCls = A.Parent.Type = vbext_ct_ClassModule
End Function

Function IsMod(A As CodeModule) As Boolean
IsMod = A.Parent.Type = vbext_ct_StdModule
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

Function MdDNm$(A As CodeModule)
MdDNm = PjnzM(A) & "." & Mdn(A)
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

Function MdFn$(A As CodeModule)
MdFn = Mdn(A) & ExtzCmpTy(CmpTyzM(A))
End Function

Function Mdn(A As CodeModule)
Mdn = A.Parent.Name
End Function

Function MdQNmzMd$(A As CodeModule)
MdQNmzMd = PjnzM(A) & "." & Mdn(A)
End Function

Function MdTy(A As CodeModule) As vbext_ComponentType
MdTy = A.Parent.Type
End Function

Function ShtCmpTyzM$(A As CodeModule)
ShtCmpTyzM = ShtCmpTy(CmpTyzM(A))
End Function

Function NUsrTyMd%(A As CodeModule)
NUsrTyMd = NUsrTySrc(DclLyzMd(A))
End Function

Function PjnzC(A As VBComponent)
PjnzC = A.Collection.Parent.Name
End Function

Function PjnzM(A As CodeModule)
PjnzM = PjnzC(A.Parent)
End Function

Function PjzM(A As CodeModule) As VBProject
Set PjzM = A.Parent.Collection.Parent
End Function

Function SizMd&(A As CodeModule)
SizMd = Len(SrcLines(A))
End Function

Function SrcLines$(A As CodeModule)
SrcLines = JnCrLf(Src(A))
End Function

Function RmvMthInSrc(Src$(), MthnSet As Aset) As String()
Dim D As Dictionary: Set D = DiceKeySet(MthnDic(Src), MthnSet): 'Brw D: Stop
RmvMthInSrc = LyzLinesDicItems(D)
End Function

Property Get CMd() As CodeModule
Set CMd = CPne.CodeModule
End Property

Property Get CMdDNm$()
CMdDNm = MdQNmzMd(CMd)
End Property

Sub ClsMd(A As CodeModule)
A.CodePane.Window.Close
End Sub

Sub CmpMdAB(A As CodeModule, B As CodeModule)
BrwCmpDicAB MthDiczMd(A), MthDiczMd(B), MdQNmzMd(A), MdQNmzMd(B)
End Sub

Sub RmvMdLno(A As CodeModule, Lno)
A.DeleteLines Lno, 1
End Sub

Private Function Y_Md() As CodeModule
Set Y_Md = CVbe.VBProjects("StockShipRate").VBComponents("Schm").CodeModule
End Function

Private Sub ZZ_MdDrs()
'BrwDrs MdDrs(Md("IdeFeature_EnsZZ_AsPrivate"))
End Sub

Private Sub ZZ_MthLnoMdMth()
Dim O$()
    Dim Lno, L&(), M, A As CodeModule, Ny$(), J%
    Set A = Md("Fct")
    Ny = MthnyzMd(A)
    For Each M In Ny
        DoEvents
        J = J + 1
        Push L, MthLnoMdMth(A, CStr(M))
        If J Mod 150 = 0 Then
            Debug.Print J, Si(Ny), "Z_MthLnoMdMth"
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
