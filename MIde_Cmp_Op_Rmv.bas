Attribute VB_Name = "MIde_Cmp_Op_Rmv"
Option Explicit
Sub DltCmpz(A As VBProject, MdNm$)
If Not HasCmpz(A, MdNm) Then Exit Sub
A.VBComponents.Remove A.VBComponents(MdNm)
End Sub
Sub RmvMdzPfx(Pfx$)
Dim Ny$(): Ny = AywPfx(MdNyPj, Pfx)
If Sz(Ny) = 0 Then InfoLin CSub, "no module begins with " & Pfx: Exit Sub
Brw Ny
Dim N
If Cfm("Rmv those Md as show in the notepad?") Then
    For Each N In Ny
        RmvMd Md(N)
    Next
End If
End Sub
Sub RmvMd(A As CodeModule)
Dim M$, P$
    M = MdNm(A)
    P = PjNmzMd(A)
'Debug.Print FmtQQ("RmvMd: Before Md(?) is deleted from Pj(?)", M, P)
A.Parent.Collection.Remove A.Parent
Debug.Print FmtQQ("RmvMd: Md(?) is deleted from Pj(?)", M, P)
End Sub

Sub RmvCmp(A As VBComponent)
A.Collection.Remove A
End Sub

