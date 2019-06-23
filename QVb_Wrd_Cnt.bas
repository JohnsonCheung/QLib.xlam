Attribute VB_Name = "QVb_Wrd_Cnt"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Wrd_Cnt."
Private Const Asm$ = "QVb"
Private Sub Z_WrdCntDic()
Dim A As Dictionary
Set A = SrtDic(WrdCntDic(JnCrLf(SrczP(CPj))))
BrwDic A
End Sub
Function WrdCntg$(S)
Dim NW&, ND&, Sy$()
Sy = WrdSy(S)
NW = Si(Sy)
ND = Si(AwDist(Sy))
WrdCntg = FmtQQ("Len: ?|Lines: ?|Words: ?|Distinct Words: ?", Len(S), NLines(S), NW, ND)
End Function
Function NWrd&(S)
NWrd = Si(WrdSy(S))
End Function
Function NDistWrd&(S)
NDistWrd = Si(AwDist(WrdSy(S)))
End Function
Function WrdCntDic(S) As Dictionary
Set WrdCntDic = CntDic(WrdSy(S))
End Function
Function WrdAset(S) As Aset
Set WrdAset = AsetzAy(WrdSy(S))
End Function

Function CvMch(A) As IMatch
Set CvMch = A
End Function

Function FstWrdAsetP() As Aset
Dim I, F$
Set FstWrdAsetP = New Aset
For Each I In SrczP(CPj)
    F = FstWrd(CStr(I))
    FstWrdAsetP.PushItm F
Next
End Function
Function FstWrd$(S)
Dim A As MatchCollection
Set A = MchWrd(S)
Select Case A.Count
Case 0: Exit Function
Case Else: FstWrd = CvMch(A(0)).Value
End Select
End Function

Function MchgWrdRe() As RegExp
Static X As RegExp
Const C$ = "[a-zA-Z][a-zA-Z0-9_]*"
If IsNothing(X) Then Set X = RegExp(C, IsGlobal:=True)
Set MchgWrdRe = X
End Function

Function MchWrd(S) As MatchCollection
Set MchWrd = MchgWrdRe.Execute(S)
End Function

Function WrdSy(S) As String()
Dim I As Match
For Each I In MchWrd(S)
    PushI WrdSy, I.Value
Next
End Function

