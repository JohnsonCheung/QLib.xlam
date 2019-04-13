Attribute VB_Name = "MIde_Identifier"
Option Explicit
Private Sub Z_NyzStr()
Dim S$
GoSub ZZ
'GoSub T0
Exit Sub
ZZ:
    Dim Lines$: Lines = SrcLinesInPj
    Dim Ny1$(): Ny1 = NyzStr(Lines)
    Dim Ny2$(): Ny2 = WrdAy(Lines)
    If Not IsEqAy(Ny1, Ny2) Then Stop
    Return
T0:
    S = "S_S"
    Ept = Sy("S_S")
    GoTo Tst
Tst:
    Act = NyzStr(S)
    C
    Return
End Sub
Private Sub Z_NsetzStr()
NsetzStr(SrcLinesInPj).Srt.Vc
End Sub
Function NsetzStr(S) As Aset
Set NsetzStr = AsetzAy(NyzStr(S))
End Function
Function RplPun$(Str)
Dim O$(), C$, A%, J&, L&
L = Len(Str)
ReDim O(L - 1)
For J = 0 To L - 1
    C = Mid(Str, J + 1, 1)
    A = Asc(C)
    If IsAscPun(A) Then O(J) = " " Else O(J) = C
Next
RplPun = Jn(O)
End Function

Function SyeNonNm(Sy$()) As String()
Dim Nm
For Each Nm In Sy
    If IsNm(Nm) Then PushI SyeNonNm, Nm
Next
End Function

Function NyzStr(S) As String()
NyzStr = SyeNonNm(SySsl(RplLf(RplCr(RplPun(S)))))
End Function
Function RelOf_PubMthNm_To_ModNy_InPj() As Rel
Set RelOf_PubMthNm_To_ModNy_InPj = RelOf_PubMthNm_To_ModNy_zPj(CurPj)
End Function
Function RelOfMthNmToCmlInVbe(Optional WhStr$) As Rel
Set RelOfMthNmToCmlInVbe = RelOfMthNmToCmlzVbe(CurVbe, WhStr)
End Function
Function RelOfMthNmToCmlzVbe(A As Vbe, Optional WhStr$) As Rel
Dim O As New Rel, MthNm
For Each MthNm In MthNyzVbe(A, WhStr)
    O.PushRelLin CmlLin(MthNm)
Next
Set RelOfMthNmToCmlzVbe = O
End Function
Function RelOf_PubMthNm_To_ModNy_zPj(A As VBProject) As Rel
Dim C, S$(), O As New Rel, MthNm, ModNm$, Cmp As VBComponent, B As WhMth
Set B = WhMthzStr("-Pub")
For Each C In CmpItr(A, "-Mod")
    Set Cmp = C
    ModNm = Cmp.Name
    S = Src(Cmp.CodeModule)
    For Each MthNm In Itr(MthNyzSrc(S, B))
        O.PushParChd MthNm, ModNm
    Next
Next
Set RelOf_PubMthNm_To_ModNy_zPj = O
End Function
Function RelOf_MthNm_To_MdNy_zPj(A As VBProject) As Rel
Dim C As VBComponent, O As New Rel, MthNm, MdNm$
For Each C In A.VBComponents
    MdNm = C.Name
    For Each MthNm In Itr(MthNyzSrc(Src(C.CodeModule)))
        O.PushParChd MthNm, MdNm
    Next
Next
Set RelOf_MthNm_To_MdNy_zPj = O
End Function
Function RelOf_MthNm_To_MdNy_InPj() As Rel
Static O As Rel
If IsNothing(O) Then Set O = RelOf_MthNm_To_MdNy_zPj(CurPj)
Set RelOf_MthNm_To_MdNy_InPj = O
End Function
Function MthExtNy(MthPjDotMdNm$, PubMthLy$(), PubMthNm_To_PjDotModNy As Dictionary) As String()
Dim Cxt$: Cxt = JnSpc(MthCxtLy(PubMthLy))
Dim Ny$(): Ny = NyzStr(Cxt)
Dim Nm
For Each Nm In Itr(Ny)
    If PubMthNm_To_PjDotModNy.Exists(Nm) Then
        Dim PjDotModNy$():
            PjDotModNy = AyeEle(PubMthNm_To_PjDotModNy(Nm), MthPjDotMdNm)
        If HasEle(PjDotModNy, Nm) Then
            PushI MthExtNy, Nm
        End If
    End If
Next
End Function

Property Get VbKwAy() As String()
Static X$()
If Si(X) = 0 Then
    X = SySsl("Function Sub Then If As For To Each End While Wend Loop Do Static Dim Option Explicit Compare Text")
End If
VbKwAy = X
End Property

Property Get VbKwAset() As Aset
Set VbKwAset = AsetzAy(VbKwAy)
End Property
