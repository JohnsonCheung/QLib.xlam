Attribute VB_Name = "MIde_Identifier"
Option Explicit
Function MthExtNy(Src$()) As String()
Dim ArgA$(): ArgA = ArgNy(ContLin(Src, 0))
Dim DimA$(): DimA = DimNy(Src)
MthExtNy = AyMinusAp(IdentifierAy(JnSpc(Src)), DimA, ArgA)
End Function
Function DimNy(Lin) As String()

End Function
Function DimNyzSrc(Src$()) As String()
Dim L
For Each L In Itr(Src)
    PushIAy DimNyzSrc, DimNy(L)
Next
End Function
Function DimNyLin(Lin) As String()

End Function

Private Sub Z_IdentifierAset()
Dim A As Aset
Set A = IdentifierAset(LinesPj(CurPj))
Debug.Print A.Cnt
A.Srt.Brw
End Sub
Function IdentifierAset(S) As Aset
Set IdentifierAset = AsetzAy(IdentifierAy(S))
End Function
Function RmvPun$(S)
If S = "" Then Exit Function
Dim J&, O$(), C$, A%
ReDim O(Len(S) - 1)
For J = 1 To Len(S)
    C = Mid(S, J, 1)
    A = Asc(C)
    If IsAscNmChr(A) Then
        O(J - 1) = C
    Else
        O(J - 1) = " "
    End If
Next
RmvPun = Jn(O, "")
End Function
Function IdentifierAy(S) As String()
Dim I
For Each I In Itr(SySsl(RmvPun(S)))
    If IsAscFstNmChr(Asc(FstChr(I))) Then
        PushI IdentifierAy, I
    End If
Next
End Function

Property Get VbKwAy() As String()
Static X$()
If Sz(X) = 0 Then
    X = SySsl("Function Sub Then If As For To Each End While Wend Loop Do Static Dim Option Explicit Compare Text")
End If
VbKwAy = X
End Property

Property Get VbKwAset() As Aset
Set VbKwAset = AsetzAy(VbKwAy)
End Property
