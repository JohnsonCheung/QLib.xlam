Attribute VB_Name = "MIde_Cur_CdPne_Md_Mth"
Option Explicit
Property Get CurMdNm$()
CurMdNm = CurCmp.Name
End Property
Property Get CurLno&()
Dim R1&, R2&, C1&, C2&
CurCdPne.GetSelection R1, C1, R2, C2
CurLno = R1
End Property
Property Get CurMthNm$()
CurMthNm = CurMthNmMd(CurMd)
End Property
Function CurMthNmMd$(A As CodeModule)
Dim K As vbext_ProcKind
CurMthNmMd = A.ProcOfLine(CurLno, K)
End Function

Property Get CurWinzMd() As VBIDE.Window
Dim A As CodePane
Set A = CurCdPne
If IsNothing(A) Then Exit Property
Set CurWinzMd = A.Window
End Property

Property Get CurCdPne() As VBIDE.CodePane
Set CurCdPne = CurVbe.ActiveCodePane
End Property
Property Get CurMthLin$()
Dim A As CodeModule: Set A = CurMd
Dim Lno&: Lno = CurLno
Dim J&
For J = Lno To 1 Step -1
    If IsMthLin(A.Lines(J, 1)) Then
        CurMthLin = ContLinzMd(A, J)
        Exit Function
    End If
Next
End Property
