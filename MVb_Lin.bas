Attribute VB_Name = "MVb_Lin"
Option Explicit

Function HasDDRmk(A) As Boolean
HasDDRmk = HasSubStr(A, "--")
End Function

Function IsSngTermLin(A) As Boolean
IsSngTermLin = InStr(Trim(A), " ") = 0
End Function

Function IsDDLin(A) As Boolean
IsDDLin = FstTwoChr(LTrim(A)) = "--"
End Function

Function IsDotLin(A) As Boolean
IsDotLin = FstChr(A) = "."
End Function

Function HasLinT1Ay(Lin, T1Ay$()) As Boolean
HasLinT1Ay = HasEle(T1Ay, T1(Lin))
End Function

Function PfxLinAp(A, ParamArray PfxAp())
Dim Av(): Av = PfxAp
Dim X
For Each X In Av
    If HasPfx(A, X) Then PfxLinAp = X: Exit Function
Next
End Function

