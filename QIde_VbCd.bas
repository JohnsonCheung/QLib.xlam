Attribute VB_Name = "QIde_VbCd"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_VbCd."
Private Const Asm$ = "QIde"
Function CdLyP() As String()
CdLyP = CdLyzS(SrczP(CPj))
End Function
Function CdLyzMd(M As CodeModule) As String()
CdLyzMd = CdLyzS(Src(M))
End Function
Function CdLyzP(P As VBProject) As String()
CdLyzP = CdLyzS(SrczP(P))
End Function
Function CdLyzS(Src$()) As String()
Dim L$, I
For Each I In Itr(Src)
    I = L
    If IsLinCd(L) Then
        PushI CdLyzS, L
    End If
Next
End Function

Function IsLinCd(Lin) As Boolean
Dim L$: L = Trim(Lin)
If Lin = "" Then Exit Function
If FstChr(LTrim(Lin)) = "'" Then Exit Function
IsLinCd = True
End Function
Function IsLinNonOpt(Lin) As Boolean
If Not IsLinCd(Lin) Then Exit Function
If HasPfx(Lin, "Option") Then Exit Function
IsLinNonOpt = True
End Function

'
