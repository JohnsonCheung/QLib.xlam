Attribute VB_Name = "MIde_VbCd"
Option Explicit
Function CdLyPj() As String()
CdLyPj = CdLyzSrc(SrcPj)
End Function
Function CdLyzMd(A As CodeModule) As String()
CdLyzMd = CdLyzSrc(Src(A))
End Function
Function CdLyzPj(A As VBProject) As String()
CdLyzPj = CdLyzSrc(SrczPj(A))
End Function
Function CdLyzSrc(Src$()) As String()
Dim L
For Each L In Itr(Src)
    If IsCdLin(L) Then
        PushI CdLyzSrc, L
    End If
Next
End Function

Function IsCdLin(A) As Boolean
Dim L$: L = Trim(A)
If A = "" Then Exit Function
If FstChr(A) = "'" Then Exit Function
IsCdLin = True
End Function
Function IsNonOptCdLin(A) As Boolean
If Not IsCdLin(A) Then Exit Function
If HasPfx(A, "Option") Then Exit Function
IsNonOptCdLin = True
End Function
