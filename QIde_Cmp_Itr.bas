Attribute VB_Name = "QIde_Cmp_Itr"
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Cmp_Itr."

Function ClsAyPj(P As VBProject, Optional WhStr$) As CodeModule()
If WhStr = "" Then
    Dim C As VBComponent
    For Each C In P.VBComponents
        If C.Type = vbext_ct_ClassModule Then
            PushObj ClsAyPj, C
        End If
    Next
Else
'    ClsAyPj = MdAyzP(A, WhStr & " -Cls")
End If
End Function

Function ClsNyPj(P As VBProject) As String()
Dim C As VBComponent
For Each C In P.VBComponents
    If IsClsCmp(C) Then PushI ClsNyPj, C.Name
Next
End Function
Private Sub Z_CmpAyzP()
Dim Act() As VBComponent
Dim C
For Each C In CmpAyzP(CPj, "-Mod")
    If CvCmp(C).Type <> vbext_ct_StdModule Then Stop
Next
For Each C In CmpAyzP(CPj, "-Cls")
    If CvCmp(C).Type <> vbext_ct_ClassModule Then Stop
Next
End Sub
Function CmpAyzP(P As VBProject, Optional WhStr$) As VBComponent()
If IsProtectzvInf(P) Then Exit Function
Dim C As VBComponent, W As WhMd
Set W = WhMdzStr(WhStr): If IsNothing(W) Then Stop '
For Each C In P.VBComponents
    If IsMd(C) Then
        If HitCmp(C, W) Then PushObj CmpAyzP, C
    End If
Next
End Function

Function IsNoClsNoModPj(P As VBProject) As Boolean
Dim C As VBComponent
For Each C In P.VBComponents
    If C.Type = vbext_ComponentType.vbext_ct_ClassModule Then Exit Function
    If C.Type = vbext_ComponentType.vbext_ct_StdModule Then Exit Function
Next
IsNoClsNoModPj = True
End Function
Function ModItrzP(P As VBProject, Optional WhStr$)
Asg Itr(ModAyzP(P, WhStr)), ModItrzP
End Function

Function ModAyzP(P As VBProject, Optional WhStr$) As CodeModule()
If P.Protection = vbext_pp_locked Then Exit Function
Dim C As VBComponent
If WhStr = "" Then
    For Each C In P.VBComponents
        PushObj ModAyzP, C.CodeModule
    Next
Else
    Dim W As WhMd: Set W = WhMdzStr(WhStr)
    For Each C In P.VBComponents
        Stop
        If HitCmp(C, W) Then
            PushObj ModAyzP, C.CodeModule
        End If
    Next
End If
End Function

Function MdNyP(Optional WhStr$) As String()
MdNyP = MdNyzP(CPj, WhStr)
End Function
Function MdNyWiPrpV(Optional WhStr$) As String()
MdNyWiPrpV = MdNyWiPrpzV(CVbe, WhStr)
End Function
Function MdNyWiPrpzV(A As Vbe, Optional WhStr$) As String()
Dim Mdn, I
For Each I In MdNyzV(A, WhStr)
    Mdn = I
    If IsMdnWiPrp(Mdn) Then
        PushI MdNyWiPrpzV, Mdn
    End If
Next
End Function

Function IsMdnWiPrp(Mdn) As Boolean
Dim M As CodeModule: Set M = Md(Mdn)
Dim J&
For J = 1 To M.CountOfLines
    If IsPrpLin(M.Lines(J, 1)) Then IsMdnWiPrp = True: Exit Function
Next
End Function

Function MdNyV(Optional WhStr$) As String()
MdNyV = MdNyzV(CVbe, WhStr)
End Function

Function MdNyzM(Mthn) As String()
MdNyzM = MdnsetzM(Mthn).Sy
End Function

Function MdAyzNy(MdNy$()) As CodeModule()
Dim N, P As VBProject
For Each N In Itr(MdNy)
    PushI MdAyzNy, MdzPN(P, N)
Next
End Function

Function MdAyzPm(PMth) As CodeModule()
MdAyzPm = MdAyzNy(MdNyzPm(PMth))
End Function

Function MdNyzPPm(P As VBProject, PMthn) As String()

End Function

Function MdNyzPm(PMthn) As String()
MdNyzPm = MdnsetzPm(PMthn).Sy
End Function

Function MdnsetzPm(PMthn) As Aset
Set MdnsetzPm = PMthnzRlMdnV.ParChd(PMthn)
End Function

Function MdnsetzM(Mthn) As Aset
Set MdnsetzM = MthnzRlMdnP.ParChd(Mthn)
End Function

Property Get PMthnzRlMdnV() As Rel
Set PMthnzRlMdnV = PMthnzRlMdnzV(CVbe)
End Property

Function PMthnzRlMdnzV(A As Vbe) As Rel
Stop
End Function

Function MdNyzP(P As VBProject, Optional WhStr$) As String()
Dim C
For Each C In CmpItr(P, WhStr)
    If IsMd(CvCmp(C)) Then
        PushI MdNyzP, C.Name
    End If
Next
End Function
Function MdNyzV(A As Vbe, WhStr$) As String()
Dim P
For Each P In PjItr(A, WhStr)
    PushIAy MdNyzV, MdNyzP(CvPj(P))
Next
End Function

Function ModNy(Optional WhStr$) As String()
ModNy = ModNyzP(CPj, WhStr)
End Function

Function ModNyzP(P As VBProject, Optional WhStr$) As String()
Dim C As VBComponent, O$()
For Each C In P.VBComponents
    If IsModCmp(C) Then PushI O, C.Name
Next
If WhStr = "" Then
    ModNyzP = O
Else
    ModNyzP = AywNmStr(O, WhStr)
End If
End Function

Private Sub Z_ClsNyPj()
DmpAy ClsNyPj(CPj)
End Sub



Private Sub Z_MdAy()
Dim O() As CodeModule
O = MdAy(CPj)
Dim I, Md As CodeModule
For Each I In O
    Set Md = I
    Debug.Print Mdn(Md)
Next
End Sub

Private Sub Z_MdzPjNy()
'DmpAy MdzPjNy(CPj)
End Sub

Private Sub ZZ()
Z_ClsNyPj
Z_MdAy
Z_MdzPjNy
MIde_Z_Pj_Cmp:
End Sub
Function CmpAy(Optional WhStr$) As VBComponent()
CmpAy = CmpAyzP(CPj, WhStr)
End Function
Function MdAy(Optional WhStr$) As CodeModule()
MdAy = MdAyzP(CPj, WhStr)
End Function

Function CmpItr(P As VBProject, Optional WhStr$)
Asg Itr(CmpAyzP(P, WhStr)), CmpItr
End Function

Function MdItr(P As VBProject, Optional WhStr$)
Asg Itr(MdAyzP(P, WhStr)), MdItr
End Function

Function MdAyzP(P As VBProject, Optional WhStr$) As CodeModule()
MdAyzP = MdAyzC(CmpAyzP(P, WhStr))
End Function

