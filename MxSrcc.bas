Attribute VB_Name = "MxSrcc"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxSrcc."
':Srcc: :Src #Src-Cleaned# ' Src without blank/rmk-Line
Function SrccP() As String()
SrccP = SrcczP(CPj)
End Function
Function SrcczM(M As CodeModule) As String()
SrcczM = Srcc(Src(M))
End Function
Function SrcczP(P As VBProject) As String()
SrcczP = Srcc(SrczP(P))
End Function

Function Srcc(Src$()) As String()
Dim L: For Each L In Itr(Src)
    If IsLinCd(L) Then PushI Srcc, L
Next
End Function

Function IsLinCd(Lin) As Boolean
Dim L$: L = Trim(Lin)
If L = "" Then Exit Function
If FstChr(L) = "'" Then Exit Function
IsLinCd = True
End Function

Function IsLinNonOpt(Lin) As Boolean
If Not IsLinCd(Lin) Then Exit Function
If HasPfx(Lin, "Option") Then Exit Function
IsLinNonOpt = True
End Function
