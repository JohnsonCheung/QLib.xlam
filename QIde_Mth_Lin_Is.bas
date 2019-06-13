Attribute VB_Name = "QIde_Mth_Lin_Is"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Lin_Is."
Private Const Asm$ = "QIde"
Function IsLinzPubMth(Lin) As Boolean
Dim L$: L = Lin
Dim Mdy$: Mdy = ShfMdy(L): If Mdy <> "" And Mdy <> "Public" Then Exit Function
IsLinzPubMth = TakMthKd(Lin) <> ""
End Function

Function IsLinzMth(Lin) As Boolean
IsLinzMth = MthKd(Lin) <> ""
End Function
Function IsMthLinzNm(Lin, Nm) As Boolean
IsMthLinzNm = Mthn(Lin) = Nm
End Function

Function MthLnozM&(M As CodeModule, Lno&)
Dim J&
For J = Lno To 1 Step -1
    If IsLinzMth(M.Lines(J, 1)) Then
        MthLnozM = J
        Exit Function
    End If
Next

End Function
Function MthLinzML$(M As CodeModule, Lno&)
MthLinzML = ContLinzML(M, MthLnozM(M, Lno))
End Function

