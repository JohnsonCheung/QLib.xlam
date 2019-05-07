Attribute VB_Name = "QApp_Fun"
Option Explicit
Private Const CMod$ = "MApp_Fun."
Private Const Asm$ = "QApp"
Private X_Acs As New Access.Application
Sub Ens()
'EnsMdCSub
'EnsMdOptExp
'EnsMdSubZZZ
Srt
End Sub

Property Get AutoExec()
'D "AutoExec:"
'D "-Before LnkCcm: CnSy--------------------------"
'D CnSy
'D "-Before LnkCcm: Srcy--------------------------"
'D Srcy
'
'EnsTblSpec

LnkCcm CurrentDb, IsDev
'D "-After LnkCcm: CnSy--------------------------"
'D CnSy
'D "-After LnkCcm: Srcy--------------------------"
'D Srcy
End Property

Function DocWsInPj() As Worksheet
Set DocWsInPj = DocWszPj(CurPj)
End Function
Function DocWszPj(A As VBProject) As Worksheet
Dim O As New Worksheet
Set O = NewWs("TermDoc")
RgzSq DocSqzPj(A), A1zWs(O)
FmtDocWs O
Set DocWszPj = O
End Function
Private Sub FmtDocWs(DocWs As Worksheet)

End Sub
Private Function DocSqzPj(A As VBProject) As Variant()
Dim Dry(), K$, I, D As Dictionary, TermAset As Aset
PushI Dry, SyzSsLin("Term Lnk DfnStmt")
Set D = DocDiczPj(A)
Set TermAset = AsetzItr(D.Keys)
For Each I In D.Keys
    K = I
    PushIAy Dry, DryzDocDfn(K, D(K), TermAset)
Next
DocSqzPj = SqzDry(Dry)
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
If IsNothing(X) Then Set X = DocDiczPj(CurPj)
Set DocDicP = X
End Function

Function DocDiczStrCnsts(A As StrCnsts) As Dictionary
Dim DocNm$, O As New Dictionary
Dim J&, M As StrCnst
A = StrCnsts(SplitCrLf(Dcl))
For J = 0 To A.N - 1
    M = A.Ay(J)
    If HasPfx(M.Nm, "Docz") Then
        DocNm = Mid(M.Nm, 6)
        O.Add DocNm, M.Str
    End If
Next
Set DocDiczStrCnsts = O
End Function
Function DocDiczDcl(Dcl$) As Dictionary
Set DocDiczDcl = DocDiczStrCnsts(StrCnsts(SplitCrLf(Dcl)))
End Function


Property Get IsDev() As Boolean
Static X As Boolean, Y As Boolean
If Not X Then
    X = True
    Y = Not HasPth(ProdPth)
End If
IsDev = Y
End Property

Property Get IsProd() As Boolean
IsProd = Not IsDev
End Property

Property Get ProdPth$()
ProdPth = "N:\SAPAccessReports\"
End Property
Property Get OHApnFb$()
OHApnFb = ApnFb("OverHeadExpense6")
End Property

Function ApnFb$(Apn$)
ApnFb = "C:\Users\Public\" & Apn & "\" & Apn & ".accdb"
End Function
Property Let ApnzDb(A As Database, V$)
ValzQ(A, SqlSel_F("Apn")) = V
End Property

Property Get ApnzDb$(A As Database)
ApnzDb = ValzQ(A, "Select Apn from Apn")
End Property

