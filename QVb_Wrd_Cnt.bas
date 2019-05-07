Attribute VB_Name = "QVb_Wrd_Cnt"
Option Explicit
Private Const CMod$ = "MVb_Wrd_Cnt."
Private Const Asm$ = "QVb"
Private Const WrdReStr$ = "[a-zA-Z][a-zA-Z0-9_]*"
Private Sub Z_WrdCntDic()
Dim A As Dictionary
Set A = DicSrt(WrdCntDic(JnCrLf(SrcInPj)))
BrwDic A
End Sub
Function WrdCntDic(S) As Dictionary
Set WrdCntDic = CntDic(WrdAy(S))
End Function
Function WrdAset(S) As Aset
Set WrdAset = AsetzAy(WrdAy(S))
End Function

Function CvMch(A) As IMatch
Set CvMch = A
End Function
Function FstWrdAsetInPjSrc() As Aset
Dim I
Set FstWrdAsetInPjSrc = New Aset
For Each I In SrcInPj
    FstWrdAsetInPjSrc.PushItm FstWrd(I)
Next
End Function
Function FstWrd$(S)
Dim A As MatchCollection
Set A = WrdMch(S)
Select Case A.Count
Case 0: Exit Function
Case 1: FstWrd = CvMch(A.Item(0)).Value
Case Else: ThwIfNEver CSub
End Select
End Function

Function WrdMch(S) As MatchCollection
Set WrdMch = RegExp(WrdReStr, IsGlobal:=True).Execute(S)
End Function

Function WrdAy(S) As String()
Dim I As Match
For Each I In WrdMch(S)
    PushI WrdAy, I.Value
Next
End Function

