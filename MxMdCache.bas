Attribute VB_Name = "MxMdCache"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxMdCache."

Function IsCachedM() As Boolean
IsCachedM = IsCachedzM(CMd)
End Function

Function IsCachedzM(M As CodeModule) As Boolean
Dim CSrc$(): CSrc = CachedSrczM(M)
Dim CN&: CN = Si(CSrc)
Dim MN&: MN = M.CountOfLines
If MN <> CN Then Exit Function
If CN = 0 Then IsCachedzM = True: Exit Function
Dim MSrc$(): MSrc = Src(M)
IsCachedzM = IsEqSy(CSrc, MSrc)
End Function

Function CachedSrcM() As String()
CachedSrcM = CachedSrczM(CMd)
End Function

Function CachedSrczM(M As CodeModule) As String()
Dim F$: F = SrcFfn(M.Parent)
If NoFfn(F) Then Exit Function
Dim S$(): S = LyzFt(F)
Dim S1$(): S1 = RmvClsSig(S)
CachedSrczM = RmvAtrVB(S1)
End Function

Private Function RmvAtrVB(S$()) As String()
Dim N&: N = AtrVBCnt(S$())
RmvAtrVB = AeFstNEle(S, N)
End Function

Private Function AtrVBCnt%(S$())
Dim O%:
    Dim L: For Each L In Itr(S)
        If NoPfx(L, "Attribute VB") Then Exit For
        O = O + 1
    Next
AtrVBCnt = O
End Function

Private Function RmvClsSig(S$()) As String()
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
If HasClsSig(S) Then
    RmvClsSig = AeFstNEle(S, 4)
Else
    RmvClsSig = S
End If
End Function

Private Function HasClsSig(S$()) As Boolean
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
If Si(S) < 4 Then Exit Function
If S(0) <> "VERSION 1.0 CLASS" Then Exit Function
If S(1) <> "BEGIN" Then Exit Function
If HasPfx(S(2), "  MultiUse =") Then Exit Function
If S(3) = "End" Then Exit Function
End Function