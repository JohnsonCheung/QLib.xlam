Attribute VB_Name = "QVbIde_Doc_DocWs"
Option Explicit
Option Compare Text

Function DocWsP() As Worksheet
Set DocWsP = DocWszP(CPj)
End Function
Function DocWszP(P As VBProject) As Worksheet
Dim O As New Worksheet
Set O = NewWs("TermDoc")
RgzSq DocSqzP(P), A1zWs(O)
FmtDocWs O
Set DocWszP = O
End Function
Private Sub FmtDocWs(DocWs As Worksheet)

End Sub
Private Function DocSqzP(P As VBProject) As Variant()
Dim Dry(), K$, I, D As Dictionary, TermAset As Aset
PushI Dry, SyzSS("Term Lnk DfnStmt")
'Set D = DocDiczP(P)
Set TermAset = AsetzItr(D.Keys)
For Each I In D.Keys
    K = I
    PushIAy Dry, DryzDocDfn(K, D(K), TermAset)
Next
DocSqzP = SqzDry(Dry)
End Function

Function DryzDocDfn(Nm$, Lines$, TermAset As Aset) As Variant()  'DocDfn = Nm + Lines.  DocDry = Term + Lnk + Stmt
Dim S$(): S = StmtLy(Lines)
Dim N0$(): N0 = AywDist(NyzStr(Lines))
Dim N1$(): N1 = AywInAset(N0, TermAset)
Dim N$(): N = AyeEle(N1, Nm)
Dim J%, Term, Nm1, Stmt, UN%, US%
UN = UB(N)
US = UB(S)
Dim O()
ReDim O(1 To Max(US, UN) + 1, 1 To 3)
O(1, 1) = Nm
For J = 0 To US
    O(J + 1, 3) = S(J)
Next
For J = 0 To UN
    O(J + 1, 2) = N(J)
Next
DryzDocDfn = DryzSq(O)
End Function

Function DocDicP() As Dictionary
Static X As Dictionary
'If IsNothing(X) Then Set X = DocDiczP(CPj)
Set DocDicP = X
End Function

