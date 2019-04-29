Attribute VB_Name = "MIde_VbCd"
Option Explicit
Function CdLyPj() As String()
CdLyPj = CdLyzSrc(SrcInPj)
End Function
Function CdLyzMd(A As CodeModule) As String()
CdLyzMd = CdLyzSrc(Src(A))
End Function
Function CdLyzPj(A As VBProject) As String()
CdLyzPj = CdLyzSrc(SrczPj(A))
End Function
Function CdLyzSrc(Src$()) As String()
Dim L$, I
For Each I In Itr(Src)
    I = L
    If IsCdLin(L) Then
        PushI CdLyzSrc, L
    End If
Next
End Function

Function IsCdLin(Lin$) As Boolean
Dim L$: L = Trim(Lin)
If Lin = "" Then Exit Function
If FstChr(LTrim(Lin)) = "'" Then Exit Function
IsCdLin = True
End Function
Function IsNonOptCdLin(Lin$) As Boolean
If Not IsCdLin(Lin) Then Exit Function
If HasPfx(Lin, "Option") Then Exit Function
IsNonOptCdLin = True
End Function
