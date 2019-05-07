Attribute VB_Name = "QIde_Ens_MthMdy_PubZ_Get"
Option Explicit
Private Const CMod$ = "MIde_Ens_MthMdy_PubZ_Get."
Private Const Asm$ = "QIde"
Private Const Ns$ = ""
Sub BrwMdNyzWiPubZ()
Brw MdNyzWiPubZ, "MdNyzWiPubZ"
End Sub
Function MdNyzWiPubZ() As String()
MdNyzWiPubZ = MdNyzWiPubZPj(CurPj)
End Function
Function MdNyzWiPubZPj(A As VBProject) As String()
Dim C As VBComponent
For Each C In A.VBComponents
    If IsWiPubZMd(C.CodeModule) Then PushI MdNyzWiPubZPj, C.Name
Next
End Function

Function MthLinSyzPubZInMd(A As CodeModule) As String()
Dim MthLin
For Each MthLin In Itr(MthLinSyzSrc(Src(A)))
    If IsMthLinzPubZ(MthLin) Then PushI MthLinSyzPubZInMd, MthLin
Next
End Function

Function MthLinSyzPubZ() As String()
MthLinSyzPubZ = MthLinSyzPubZInPj(CurPj)
End Function
Function MthLinSyzPubZInPj(A As VBProject) As String()
If A.Protection = vbext_pp_locked Then Exit Function
Dim C As VBComponent
For Each C In A.VBComponents
    PushIAy MthLinSyzPubZInPj, MthLinSyzPubZInMd(C.CodeModule)
Next
End Function
Function IsWiPubZMd(A As CodeModule) As Boolean
Dim MthLin
For Each MthLin In Itr(MthLinSyzSrc(Src(A)))
    If IsMthLinzPubZ(MthLin) Then IsWiPubZMd = True: Exit Function
Next
End Function

Function MthLinSyzPub(Src$()) As String()
Dim L
For Each L In Itr(Src)
    If IsMthLinzPub(L) Then
        PushI MthLinSyzPub, L
    End If
Next
End Function


