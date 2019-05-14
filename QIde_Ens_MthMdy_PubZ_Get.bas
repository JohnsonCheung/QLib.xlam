Attribute VB_Name = "QIde_Ens_MthMdy_PubZ_Get"
Option Explicit
Private Const CMod$ = "MIde_Ens_MthMdy_PubZ_Get."
Private Const Asm$ = "QIde"
Private Const NS$ = ""
Sub BrwMdNyzWiPubZ()
Brw MdNyzWiPubZ, "MdNyzWiPubZ"
End Sub
Function MdNyzWiPubZ() As String()
MdNyzWiPubZ = MdNyzWiPubzP(CPj)
End Function
Function MdNyzWiPubzP(P As VBProject) As String()
Dim C As VBComponent
For Each C In P.VBComponents
    If IsWiPubZMd(C.CodeModule) Then PushI MdNyzWiPubzP, C.Name
Next
End Function

Function MthLinyzPubZM(A As CodeModule) As String()
Dim MthLin
For Each MthLin In Itr(MthLinyzSrc(Src(A)))
    If IsMthLinzPubZ(MthLin) Then PushI MthLinyzPubZM, MthLin
Next
End Function

Function MthLinyzPubZ() As String()
MthLinyzPubZ = MthLinyzPubZP(CPj)
End Function
Function MthLinyzPubZP(P As VBProject) As String()
If P.Protection = vbext_pp_locked Then Exit Function
Dim C As VBComponent
For Each C In P.VBComponents
    PushIAy MthLinyzPubZP, MthLinyzPubZM(C.CodeModule)
Next
End Function
Function IsWiPubZMd(A As CodeModule) As Boolean
Dim MthLin
For Each MthLin In Itr(MthLinyzSrc(Src(A)))
    If IsMthLinzPubZ(MthLin) Then IsWiPubZMd = True: Exit Function
Next
End Function

Function MthLinyzPub(Src$()) As String()
Dim L
For Each L In Itr(Src)
    If IsMthLinzPub(L) Then
        PushI MthLinyzPub, L
    End If
Next
End Function


