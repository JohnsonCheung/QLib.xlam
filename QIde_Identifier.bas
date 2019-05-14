Attribute VB_Name = "QIde_Identifier"
Option Explicit
Private Const CMod$ = "MIde_Identifier."
Private Const Asm$ = "QIde"
Private Sub Z_NyzStr()
Dim S$
GoSub ZZ
'GoSub T0
Exit Sub
ZZ:
    Dim Lines$: Lines = SrcLinesP
    Dim Ny1$(): Ny1 = NyzStr(Lines)
    Dim Ny2$(): Ny2 = WrdSy(Lines)
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
NsetzStr(SrcLinesP).Srt.Vc
End Sub
Function NsetzStr(S) As Aset
Set NsetzStr = AsetzAy(NyzStr(S))
End Function
Function RplPun$(S)
Dim O$(), C$, A%, J&, L&
L = Len(S)
ReDim O(L - 1)
For J = 0 To L - 1
    C = Mid(S, J + 1, 1)
    A = Asc(C)
    If IsAscPun(A) Then O(J) = " " Else O(J) = C
Next
RplPun = Jn(O)
End Function

Function SyeNonNm(Sy$()) As String()
Dim Nm$, I
For Each I In Sy
    Nm = I
    If IsNm(Nm) Then PushI SyeNonNm, Nm
Next
End Function

Function NyzStr(S) As String()
NyzStr = SyeNonNm(SyzSS(RplLf(RplCr(RplPun(S)))))
End Function

Function PMthnzRlModNyP() As Rel
Set PMthnzRlModNyP = PMthzRlModNyzP(CPj)
End Function
Function MthnzRlCmlV(Optional WhStr$) As Rel
Set MthnzRlCmlV = MthnzRlCmlzV(CVbe, WhStr)
End Function
Function MthnzRlCmlzV(A As Vbe, Optional WhStr$) As Rel
Dim O As New Rel, I
For Each I In MthNyzV(A, WhStr)
    O.PushRelLin CmlLin(I)
Next
Set MthnzRlCmlzV = O
End Function
Function PMthzRlModNyzP(P As VBProject) As Rel
Dim C, S$(), O As New Rel, Mthn, Modn, Cmp As VBComponent, B As WhMth
Set B = WhMthzStr("-Pub")
For Each C In CmpItr(P, "-Mod")
    Set Cmp = C
    Modn = Cmp.Name
    S = Src(Cmp.CodeModule)
    For Each Mthn In Itr(MthnyzSrc(S, B))
        O.PushParChd Mthn, Modn
    Next
Next
Set PMthzRlModNyzP = O
End Function
Function MthnzRlMdnzP(P As VBProject) As Rel
Dim C As VBComponent, O As New Rel, Mthn, Mdn
For Each C In P.VBComponents
    Mdn = C.Name
    For Each Mthn In Itr(MthnyzSrc(Src(C.CodeModule)))
        O.PushParChd Mthn, Mdn
    Next
Next
Set MthnzRlMdnzP = O
End Function
Function MthnzRlMdnP() As Rel
Static O As Rel
If IsNothing(O) Then Set O = MthnzRlMdnzP(CPj)
Set MthnzRlMdnP = O
End Function
Function MthExtNy(MthPjDotMdn, PMthLy$(), PMthn_To_PjDotModNy As Dictionary) As String()
Dim Cxt$: Cxt = JnSpc(MthCxtLy(PMthLy))
Dim Ny$(): Ny = NyzStr(Cxt)
Dim Nm
For Each Nm In Itr(Ny)
    If PMthn_To_PjDotModNy.Exists(Nm) Then
        Dim PjDotModNy$():
            PjDotModNy = AyeEle(PMthn_To_PjDotModNy(Nm), MthPjDotMdn)
        If HasEle(PjDotModNy, Nm) Then
            PushI MthExtNy, Nm
        End If
    End If
Next
End Function

Property Get VbKwAy() As String()
Static X$()
If Si(X) = 0 Then
    X = SyzSS("Function Sub Then If As For To Each End While Wend Loop Do Static Dim Option Explicit Compare Text")
End If
VbKwAy = X
End Property

Property Get VbKwAset() As Aset
Set VbKwAset = AsetzAy(VbKwAy)
End Property