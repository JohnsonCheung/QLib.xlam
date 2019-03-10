Attribute VB_Name = "MIde_Dcl_EnmAndTy"
Option Explicit
Const CMod$ = "MIde_Dcl_EnmAndTy."

Function EnmBdyLySrc(Src$(), EnmNm$) As String()
EnmBdyLySrc = EnmBdyLy(EnmLy(Src, EnmNm$))
End Function
Function EnmBdyLy(EnmLy$()) As String()

End Function
Function EnmFTIx(Src$(), EnmNm) As FTIx
Dim Fm&: Fm = EnmFmIx(Src, EnmNm)
Set EnmFTIx = FTIx(Fm, EndEnmIx(Src, Fm))
End Function
Function EnmLy(Src$(), EnmNm$) As String()
EnmLy = AywFTIx(Src, EnmFTIx(Src, EnmNm))
End Function
Function EnmFmIx&(Src$(), EnmNm)
Dim J&, L
For Each L In Itr(Src)
    L = RmvMdy(L)
    If ShfXEnm(L) Then
        If TakNm(L) = EnmNm Then
            EnmFmIx = J
            Exit Function
        End If
    End If
    If IsMthLin(L) Then Exit For
    J = J + 1
Next
EnmFmIx = -1
End Function
Function EnmNyMd(A As CodeModule) As String()
EnmNyMd = EnmNy(DclLyMd(A))
End Function
Function EnmNyPj(Pj As VBProject, Optional WhStr$) As String()
Dim M
For Each M In MdItr(Pj, WhStr)
    PushIAy EnmNyPj, EnmNyMd(CvMd(M))
Next
End Function
Function EnmNy(Src$()) As String()
Dim L
For Each L In Itr(Src)
   PushNonBlankStr EnmNy, EnmNm(L)
Next
End Function

Function HasUsrTyNm(Src$(), Nm$) As Boolean
Dim L
For Each L In Itr(Src)
    If UsrTyNm(L) = Nm Then HasUsrTyNm = True: Exit Function
Next
End Function

Function NEnm%(Src$())
Dim L, O%
For Each L In Itr(Src)
   If IsEmnLin(L) Then O = O + 1
Next
NEnm = O
End Function

Function UsrTyFTIx(Src$(), TyNm$) As FTIx
Dim FmI&: FmI = UsrTyFmIx(Src, TyNm)
Dim ToI&: ToI = EndTyIx(Src, FmI)
Set UsrTyFTIx = FTIx(FmI, ToI)
End Function

Function EndEnmIx&(Src$(), FmIx)
EndEnmIx = EndLinIx(Src, "Enum", FmIx)
End Function

Function EndTyIx&(Src$(), FmIx)
EndTyIx = EndLinIx(Src, "Type", FmIx)
End Function

Function UsrTyLines$(Src$(), UsrTyNm$)
UsrTyLines = JnCrLf(UsrTyLy(Src, UsrTyNm))
End Function

Function UsrTyLy(Src$(), TyNm$) As String()
UsrTyLy = AywFTIx(Src, UsrTyFTIx(Src, TyNm))
End Function

Function UsrTyFmIx&(Src$(), TyNm)
Dim J%
For J = 0 To UB(Src)
   If IsUsrTyLin(Src(J)) = TyNm Then UsrTyFmIx = J: Exit Function
   If IsMthLin(Src(J)) Then Exit For
Next
UsrTyFmIx = -1
End Function

Function UsrTyNy(Src$()) As String()
Dim L
For Each L In Itr(Src)
    PushNonBlankStr UsrTyNy, UsrTyNm(L)
    If IsMthLin(L) Then Exit Function
Next
End Function

Function IsEmnLin(A) As Boolean
IsEmnLin = HasPfx(RmvMdy(A), "Enum ")
End Function

Function IsUsrTyLin(A) As Boolean
IsUsrTyLin = HasPfx(RmvMdy(A), "Type ")
End Function

Function EnmNm$(Lin)
Dim L$: L = RmvMdy(Lin)
If ShfX(L, "Enum ") Then EnmNm = TakNm(LTrim(L))
End Function

Function UsrTyNm$(Lin)
Dim L$: L = RmvMdy(Lin)
If ShfX(L, "Type ") Then UsrTyNm = TakNm(LTrim(L))
End Function

Function EnmLyMd(Md As CodeModule, EnmNm$) As String()
EnmLyMd = EnmLy(DclLyMd(Md), EnmNm)
End Function

Function NEnmMbrMd%(A As CodeModule, EnmNm$)
NEnmMbrMd = Sz(EnmMbrLyMd(A, EnmNm))
End Function

Function EnmMbrLyMd(A As CodeModule, EnmNm$) As String()
EnmMbrLyMd = CdLyzSrc(EnmLyMd(A, EnmNm))
End Function

Function NEnmMd%(A As CodeModule)
NEnmMd = NEnm(DclLyMd(A))
End Function

Function UsrTyNyMd(A As CodeModule) As String()
UsrTyNyMd = AySrt(UsrTyNy(DclLyMd(A)))
End Function

Function UsrTyNyPj(A As VBProject, Optional WhStr$) As String()
Dim I, M As CodeModule, O$(), W As WhNm
Set W = WhNmzStr(WhStr)
For Each I In MdItr(A, WhStr)
    Set M = CvMd(I)
    O = UsrTyNy(Src(M))
    O = AywNm(O, W)
    PushIAy UsrTyNyPj, AyAddPfx(O, MdNm(M) & ".")
Next
UsrTyNyPj = AyQSrt(O)
End Function

Function ShfXEnm(O) As Boolean
ShfXEnm = ShfX(O, "Enum")
End Function

Function ShfXTy(O) As Boolean
ShfXTy = ShfX(O, "Type")
End Function

Private Sub Z()
MIde_Dcl_EnmAndTy:
End Sub

Private Sub Z_NEnmMbrMd()
Ass NEnmMbrMd(Md("Ide"), "AA") = 1
End Sub

