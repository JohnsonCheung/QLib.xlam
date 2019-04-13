Attribute VB_Name = "MIde_Md"
Option Explicit
Const CMod$ = "MIde_Md."
Public Const DoczMdDic$ = "It is from Pj. Key is MdNm and Val is MdLines"

Function IsCls(A As CodeModule) As Boolean
IsCls = A.Parent.Type = vbext_ct_ClassModule
End Function

Function IsMod(A As CodeModule) As Boolean
IsMod = A.Parent.Type = vbext_ct_StdModule
End Function

Function Md(MdDNm) As CodeModule
Const CSub$ = CMod & "Md"
Dim A1$(): A1 = Split(MdDNm, ".")
Select Case Si(A1)
Case 1: Set Md = CurPj.VBComponents(A1(0)).CodeModule
Case 2: Set Md = Pj(A1(0)).VBComponents(A1(1)).CodeModule
Case Else: Thw CSub, "[MdDNm] should be XXX.XXX or XXX", MdDNm
End Select
End Function

Function MdAyCmpAy(A() As CodeModule) As VBComponent()
MdAyCmpAy = IntoOyP(A, "Parent", MdAyCmpAy)
End Function

Function MdAywInTy(A() As CodeModule, WhInTyAy0$) As CodeModule()
Dim TyAy() As vbext_ComponentType, Md
'TyAy = CvWhCmpTy(WhInTyAy0)
Dim O() As CodeModule
For Each Md In A
    If HasEle(TyAy, CvMd(Md).Parent.Type) Then PushObj O, Md
Next
MdAywInTy = O
End Function

Function MdDNm$(A As CodeModule)
MdDNm = PjNmzMd(A) & "." & MdNm(A)
End Function

Function MdDic(A As VBProject) As Dictionary
Dim C As VBComponent
Set MdDic = New Dictionary
For Each C In A.VBComponents
    MdDic.Add C.Name, SrcLines(C.CodeModule)
Next
End Function

Function MdDicInPj() As Dictionary
Set MdDicInPj = MdDic(CurPj)
End Function

Function MdFn$(A As CodeModule)
MdFn = MdNm(A) & SrcExtMd(A)
End Function

Function MdNm$(A As CodeModule)
MdNm = A.Parent.Name
End Function

Function MdQNmzMd$(A As CodeModule)
MdQNmzMd = PjNmzMd(A) & "." & MdNm(A)
End Function

Function MdTy(A As CodeModule) As vbext_ComponentType
MdTy = A.Parent.Type
End Function

Function MdTyNm$(A As CodeModule)
MdTyNm = ShtCmpTy(CmpTy(A))
End Function

Function NUsrTyMd%(A As CodeModule)
NUsrTyMd = NUsrTySrc(DclLyzMd(A))
End Function

Function PjNmzCmp(A As VBComponent)
PjNmzCmp = A.Collection.Parent.Name
End Function

Function PjNmzMd(A As CodeModule)
PjNmzMd = PjNmzCmp(A.Parent)
End Function

Function PjzMd(A As CodeModule) As VBProject
Set PjzMd = A.Parent.Collection.Parent
End Function

Function SizMd&(A As CodeModule)
SizMd = Len(SrcLines(A))
End Function

Function SrcLines$(A As CodeModule)
SrcLines = JnCrLf(Src(A))
End Function

Function SrcRmvMth(Src$(), MthNmSet As Aset) As String()
Dim D As Dictionary: Set D = DicExlKeySet(MthNmDic(Src), MthNmSet): 'Brw D: Stop
SrcRmvMth = LyzLinesDicItems(D)
End Function

Property Get CurMd() As CodeModule
Set CurMd = CurCdPne.CodeModule
End Property

Property Get CurMdDNm$()
CurMdDNm = MdQNmzMd(CurMd)
End Property

Sub ClsMd(A As CodeModule)
A.CodePane.Window.Close
End Sub

Sub CmpMdAB(A As CodeModule, B As CodeModule)
BrwCmpDicAB MthDiczMd(A), MthDiczMd(B), MdQNmzMd(A), MdQNmzMd(B)
End Sub

Sub RmvMdLno(A As CodeModule, Lno&)
A.DeleteLines Lno, 1
End Sub

Private Property Get ZZMd() As CodeModule
Set ZZMd = CurVbe.VBProjects("StockShipRate").VBComponents("Schm").CodeModule
End Property

Private Sub ZZ_MdDrs()
'BrwDrs MdDrs(Md("IdeFeature_EnsZZ_AsPrivate"))
End Sub

Private Sub ZZ_MthLnoMdMth()
Dim O$()
    Dim Lno, L&(), M, A As CodeModule, Ny$(), J%
    Set A = Md("Fct")
    Ny = MthNyzMd(A)
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

Private Sub Z_CurMd()
Ass CurMd.Parent.Name = "Cur_d"
End Sub
