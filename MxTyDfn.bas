Attribute VB_Name = "MxTyDfn"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxTyDfn."
Public Const FFoTyDfn$ = "Mdn Nm Ty Mem Rmk"

Private Sub Z_DoTyDfnP()
BrwDrs DoTyDfn
End Sub

Function TyDfnNyP() As String()
TyDfnNyP = TyDfnNyzP(CPj)
End Function

Function TyDfnNyzP(P As VBProject) As String()
Dim L: For Each L In VbRmk(SrczP(P))
    PushNB TyDfnNyzP, TyDfnNm(L)
Next
End Function

Private Function IsLinOkTyDfn(L) As Boolean
Dim NM$, Dfn$, T3$, Rst$
Asg3TRst L, NM, Dfn, T3, Rst
IsLinOkTyDfn = IsTyDfn(NM, Dfn, T3, Rst)
End Function

Function IsLinTyDfn(L) As Boolean
If FstChr(L) <> "'" Then Exit Function
Dim T$: T = T1(L)
If Fst2Chr(T) <> "':" Then Exit Function
If LasChr(T) <> ":" Then Exit Function
IsLinTyDfn = True
End Function

Function IsLinNkTyDfn(L) As Boolean
IsLinNkTyDfn = Not IsLinOkTyDfn(L)
End Function

Function NkTyDfnLy(Src$()) As String()
Dim L: For Each L In Itr(Src)
    If IsLinTyDfn(L) Then
        If IsLinNkTyDfn(L) Then
            PushI NkTyDfnLy, L
        End If
    End If
Next
End Function

Function TyDfnNm$(Lin)
Dim T$: T = T1(Lin)
If T = "" Then Exit Function
If Fst2Chr(T) <> "':" Then Exit Function
If LasChr(T) <> ":" Then Exit Function
TyDfnNm = RmvFstChr(T)
End Function

Function IsLinTyDfnRmk(Lin) As Boolean
If FstChr(Lin) <> "'" Then Exit Function
If FstChr(LTrim(RmvFstChr(Lin))) <> "!" Then Exit Function
IsLinTyDfnRmk = True
End Function

Function IsTyDfn(NM$, Dfn$, ThirdTerm$, Rst$) As Boolean
Select Case True
Case Fst2Chr(NM) <> "':"
Case LasChr(NM) <> ":"
Case FstChr(Dfn) <> ":"
Case ThirdTerm <> "" And Not HasPfxSfx(ThirdTerm, "#", "#") And FstChr(ThirdTerm) <> "!"
Case Else: IsTyDfn = True
End Select
End Function