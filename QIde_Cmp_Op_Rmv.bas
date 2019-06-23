Attribute VB_Name = "QIde_Cmp_Op_Rmv"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Cmp_Op_Rmv."
Private Const Asm$ = "QIde"
Sub DltCmpzPjn(P As VBProject, Mdn)
If Not HasCmpzP(P, Mdn) Then Exit Sub
P.VBComponents.Remove P.VBComponents(Mdn)
End Sub
Sub RmvMdzPfx(Pfx$)
Dim Ny$(): Ny = AwPfx(MdNyP, Pfx)
If Si(Ny) = 0 Then InfLin CSub, "no module begins with " & Pfx: Exit Sub
Brw Ny
Dim N
If Cfm("Rmv those Md as show in the notepad?") Then
    For Each N In Ny
        RmvMd Md(N)
    Next
End If
End Sub
Sub RmvMd(MdDNm)
RmvMdzMd Md(MdDNm)
End Sub
Sub RmvMdzMd(M As CodeModule)
Dim N$, P$
    N = Mdn(M)
    P = PjnzM(M)
'Debug.Print FmtQQ("RmvMd: Before Md(?) is deleted from Pj(?)", M, P)
M.Parent.Collection.Remove M.Parent
Debug.Print FmtQQ("RmvMd: Md(?) is deleted from Pj(?)", N, P)
End Sub
Sub RmvCmpzN(Cmpn)
RmvCmp Cmp(Cmpn)
End Sub
Sub RmvCmp(A As VBComponent)
A.Collection.Remove A
End Sub

