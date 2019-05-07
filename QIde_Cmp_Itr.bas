Attribute VB_Name = "QIde_Cmp_Itr"
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Cmp_Itr."

Function ClsAyPj(A As VBProject, Optional WhStr$) As CodeModule()
If WhStr = "" Then
    Dim C As VBComponent
    For Each C In A.VBComponents
        If C.Type = vbext_ct_ClassModule Then
            PushObj ClsAyPj, C
        End If
    Next
Else
'    ClsAyPj = MdAyzPj(A, WhStr & " -Cls")
End If
End Function

Function ClsNyPj(A As VBProject) As String()
Dim C As VBComponent
For Each C In A.VBComponents
    If IsClsCmp(C) Then PushI ClsNyPj, C.Name
Next
End Function
Private Sub Z_CmpAyzPj()
Dim Act() As VBComponent
Dim C
For Each C In CmpAyzPj(CurPj, "-Mod")
    If CvCmp(C).Type <> vbext_ct_StdModule Then Stop
Next
For Each C In CmpAyzPj(CurPj, "-Cls")
    If CvCmp(C).Type <> vbext_ct_ClassModule Then Stop
Next
End Sub
Function CmpAyzPj(A As VBProject, Optional WhStr$) As VBComponent()
If IsProtectzInf(A) Then Exit Function
Dim C As VBComponent, W As WhMd
Set W = WhMdzStr(WhStr): If IsNothing(W) Then Stop '
For Each C In A.VBComponents
    If IsMd(C) Then
        If HitCmp(C, W) Then PushObj CmpAyzPj, C
    End If
Next
End Function

Function MdAyzCmp(CmpAy() As VBComponent) As CodeModule()
Dim I
For Each I In Itr(CmpAy)
    PushObj MdAyzCmp, CvCmp(I).CodeModule
Next
End Function


Function IsNoClsNoModPj(A As VBProject) As Boolean
Dim C As VBComponent
For Each C In A.VBComponents
    If C.Type = vbext_ComponentType.vbext_ct_ClassModule Then Exit Function
    If C.Type = vbext_ComponentType.vbext_ct_StdModule Then Exit Function
Next
IsNoClsNoModPj = True
End Function
Function ModItrzPj(A As VBProject, Optional WhStr$)
Asg Itr(ModAyzPj(A, WhStr)), ModItrzPj
End Function
Function ModAyzPj(A As VBProject, Optional WhStr$) As CodeModule()
If A.Protection = vbext_pp_locked Then Exit Function
Dim C As VBComponent
If WhStr = "" Then
    For Each C In A.VBComponents
        PushObj ModAyzPj, C.CodeModule
    Next
Else
    Dim W As WhMd: Set W = WhMdzStr(WhStr)
    For Each C In A.VBComponents
        If HitCmp(C, W) Then
            PushObj ModAyzPj, C.CodeModule
        End If
    Next
End If
End Function

Function MdNyInPj(Optional WhStr$) As String()
MdNyInPj = MdNyzPj(CurPj, WhStr)
End Function
Function MdNyWiPrpInVbe(Optional WhStr$) As String()
MdNyWiPrpInVbe = MdNyWiPrpzVbe(CurVbe, WhStr)
End Function
Function MdNyWiPrpzVbe(A As Vbe, Optional WhStr$) As String()
Dim MdNm
For Each MdNm In MdNyzVbe(A, WhStr)
    If IsMdWiPrp(MdNm) Then
        PushI MdNyWiPrpzVbe, MdNm
    End If
Next
End Function
Function IsMdWiPrp(MdNm) As Boolean
Dim M As CodeModule: Set M = Md(MdNm)
Dim J&
For J = 1 To M.CountOfLines
    If IsPrpLin(M.Lines(J, 1)) Then IsMdWiPrp = True: Exit Function
Next
End Function
Function MdNyInVbe(Optional WhStr$) As String()
MdNyInVbe = MdNyzVbe(CurVbe, WhStr)
End Function
Function MdNyzMth(MthNm$) As String()
MdNyzMth = MdNsetzMth(MthNm).Sy
End Function
Function MdNsetzMth(MthNm$) As Aset
Set MdNsetzMth = RelOf_MthNm_To_MdNy_zPj.ParChd(MthNm)
End Function
Function MdNyzPj(A As VBProject, Optional WhStr$) As String()
Dim C
For Each C In CmpItr(A, WhStr)
    If IsMd(CvCmp(C)) Then
        PushI MdNyzPj, C.Name
    End If
Next
End Function
Function MdNyzVbe(A As Vbe, WhStr$) As String()
Dim P
For Each P In PjItr(A, WhStr)
    PushIAy MdNyzVbe, MdNyzPj(CvPj(P))
Next
End Function

Function ModNy(Optional WhStr$) As String()
ModNy = ModNyzPj(CurPj, WhStr)
End Function

Function ModNyzPj(A As VBProject, Optional WhStr$) As String()
Dim C As VBComponent, O$()
For Each C In A.VBComponents
    If IsModCmp(C) Then PushI O, C.Name
Next
If WhStr = "" Then
    ModNyzPj = O
Else
    ModNyzPj = AywNmStr(O, WhStr)
End If
End Function

Private Sub Z_ClsNyPj()
DmpAy ClsNyPj(CurPj)
End Sub



Private Sub Z_MdAy()
Dim O() As CodeModule
O = MdAy(CurPj)
Dim I, Md As CodeModule
For Each I In O
    Set Md = I
    Debug.Print MdNm(Md)
Next
End Sub

Private Sub Z_MdzPjNy()
'DmpAy MdzPjNy(CurPj)
End Sub

Private Sub Z()
Z_ClsNyPj
Z_MdAy
Z_MdzPjNy
MIde_Z_Pj_Cmp:
End Sub
Function CmpAy(Optional WhStr$) As VBComponent()
CmpAy = CmpAyzPj(CurPj, WhStr)
End Function
Function MdAy(Optional WhStr$) As CodeModule()
MdAy = MdAyzPj(CurPj, WhStr)
End Function

Function CmpItr(A As VBProject, Optional WhStr$)
Asg Itr(CmpAyzPj(A, WhStr)), CmpItr
End Function

Function MdItr(A As VBProject, Optional WhStr$)
Asg Itr(MdAyzPj(A, WhStr)), MdItr
End Function

Function MdAyzPj(A As VBProject, Optional WhStr$) As CodeModule()
MdAyzPj = MdAyzCmp(CmpAyzPj(A, WhStr))
End Function

