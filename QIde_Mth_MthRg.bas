Attribute VB_Name = "QIde_Mth_MthRg"
Type MthRg
    Mthn As String
    FmIx As Long
    EIx As Long
End Type
Type MthRgs:   N As Integer: Ay() As MthRg:   End Type
    
Function MthRg(Mthn, FmIx, EIx) As MthRg
With MthRg
    .Mthn = Mthn
    .FmIx = FmIx
    .EIx = EIx
End With
End Function

Sub PushMthRg(O As MthRgs, M As MthRg)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub

Function MthRgs(Src$()) As MthRgs
If Si(Src) = 0 Then Exit Function
Dim F&(), T&(), N$()
F = MthIxy(Src)
N = MthnyzSI(Src, F)
T = MthEIxy(Src, F)
Dim S&
S = Si(F)
If S = 0 Then Exit Function
Dim J&
For J = 0 To S - 1
    PushMthRg MthRgs, MthRg(N(J), F(J), T(J))
Next
End Function
Sub BrwMthRgs(A As MthRgs)
B LyzMthRgs(A)
End Sub
Function LyzMthRgs(A As MthRgs) As String()
Dim J%
For J = 0 To A.N - 1
    PushI LyzMthRgs, LinzMthRg(A.Ay(J))
Next
End Function
Function LinzMthRg(A As MthRg)
With A: LinzMthRg = JnSpc(Sy(.Mthn, .FmIx, .EIx)): End With
End Function

Private Sub ZZ_MthRgs()
Dim Src$()
GoSub ZZ
Exit Sub
ZZ:
    BrwMthRgs MthRgs(SrcV)
    Return
End Sub
Function MthLy(Src$(), A As MthRg) As String()
MthLy = AywFT(Src, A.FmIx, A.EIx)
End Function

