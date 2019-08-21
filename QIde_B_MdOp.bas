Attribute VB_Name = "QIde_B_MdOp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Md_Op_Rmv_Lines."
Private Const Asm$ = "QIde"

Sub ClrMd(M As CodeModule)
With M
    If .CountOfLines = 0 Then Exit Sub
    Debug.Print FmtQQ("ClrMd: Md(?) of Lines(?) is cleared", Mdn(M), .CountOfLines)
    .DeleteLines 1, .CountOfLines
    If .CountOfLines <> 0 Then Stop
End With
End Sub

Function CntSiStrzMd$(M As CodeModule)
CntSiStrzMd = CntSiStrzLines(SrcL(M))
End Function

Sub DltLin(M As CodeModule, Lno, OldLin)
Dim LinFmMd$: LinFmMd = M.Lines(Lno, 1)
If LinFmMd <> OldLin Then Thw CSub, "Lines from Md <> OldLines", "Md Lno Lines-from-Md OldLines", Mdn(M), Lno, LinFmMd, OldLin
M.DeleteLines Lno, 1
End Sub

Sub DltLines(M As CodeModule, Lno&, OldLines$)
If OldLines = "" Then Exit Sub
If Lno = 0 Then Exit Sub
Dim Cnt&: Cnt = LinCnt(OldLines)
If M.Lines(Lno, Cnt) <> OldLines Then Thw CSub, "OldL <> ActL", "OldL ActL", OldLines, M.Lines(Lno, Cnt)
Debug.Print FmtQQ("DltLines: Lno(?) Cnt(?)", Lno, Cnt)
D Box(SplitCrLf(OldLines))
D ""
M.DeleteLines Lno, Cnt
End Sub

Sub DltLinzF(M As CodeModule, B As Feis)
If Not IsFeisInOrd(B) Then Thw CSub, "Given Feis is not in order", "Feis", LyzFeis(B)
Dim J%
For J = B.N - 1 To 0 Step -1
    With FCntzFei(B.Ay(J))
        M.DeleteLines .FmLno, .Cnt
    End With
Next
End Sub

Sub DltLinzFei(M As CodeModule, B As Fei, OldLines$)
Stop
Dim FstLin
'FstLin = A.Lines(Fei.FmNo, 1)
With B
'    If .Cnt = 0 Then Exit Sub
'    A.DeleteLines .FmNo, .Cnt
End With
End Sub

Sub DltLinzFeis(M As CodeModule, B As Feis)
If Not IsFeisInOrd(B) Then Stop
Dim J&
For J = B.N - 1 To 0 Step -1
'    DltLinzFEITx B.Ay(J)
Next
End Sub

Sub InsLines(M As CodeModule, Lno, Lines$)
M.InsertLines Lno, Lines
End Sub

Function LinzFei$(A As Fei)
With A
LinzFei = "FmEndIx " & .FmIx & " " & .EIx
End With
End Function

Function LyzFeis(A As Feis) As String()
Dim J&
For J = 0 To A.N - 1
    PushI LyzFeis, J & " " & LinzFei(A.Ay(J))
Next
End Function

Sub DltLinzD(M As CodeModule, L_OldL As Drs)
If JnSpc(L_OldL.Fny) <> "L OldL" Then Stop: Exit Sub
Stop
Dim B As Drs: B = SrtDrs(L_OldL, "-L")
Dim Dr
Stop
For Each Dr In Itr(B.Dy)
    Dim L&: L = Dr(0)
    Dim OldL$: OldL = Dr(2)
    Dim NewL$: NewL = Dr(1)
    If M.Lines(L, 1) <> OldL Then Thw CSub, "Md-Lin <> OldL", "Mdn Lno Md-Lin OldL NewL", Mdn(M), L, M.Lines(L, 1), OldL, NewL
    M.ReplaceLine L, NewL
Next
End Sub

Sub InsLinzD(M As CodeModule, L_NewL As Drs)
If JnSpc(L_NewL.Fny) <> "L NewL" Then Stop: Exit Sub
Stop
Dim B As Drs: B = SrtDrs(L_NewL, "-L")
Dim Dr
Stop
For Each Dr In Itr(B.Dy)
    Dim L&: L = Dr(0)
    Dim OldL$: OldL = Dr(2)
    Dim NewL$: NewL = Dr(1)
    If M.Lines(L, 1) <> OldL Then Thw CSub, "Md-Lin <> OldL", "Mdn Lno Md-Lin OldL NewL", Mdn(M), L, M.Lines(L, 1), OldL, NewL
    M.ReplaceLine L, NewL
Next
End Sub
Sub InsLinzDcl(M As CodeModule, Ly$())
Dim Lno&: Lno = LnozFstDcl(M)
Dim L: For Each L In Itr(Ly)
    M.InsertLines Lno, L
Next
End Sub
Sub InsLin(M As CodeModule, L_NewL As Drs)
Dim B As Drs: B = L_NewL
If JnSpc(B.Fny) <> "L NewL" Then Stop: Exit Sub
Dim Dr
For Each Dr In Itr(B.Dy)
    Dim L&: L = Dr(0)
    Dim NewL$: NewL = Dr(1)
    M.InsertLines L, NewL
Next
End Sub

Sub RplLin(M As CodeModule, L_NewL_OldL As Drs)
Dim B As Drs: B = L_NewL_OldL
If JnSpc(B.Fny) <> "L NewL OldL" Then Stop: Exit Sub
Dim Dr
For Each Dr In Itr(B.Dy)
    Dim L&: L = Dr(0)
    Dim OldL$: OldL = Dr(2)
    Dim NewL$: NewL = Dr(1)
    If M.Lines(L, 1) <> OldL Then Thw CSub, "Md-Lin <> OldL", "Mdn Lno Md-Lin OldL NewL", Mdn(M), L, M.Lines(L, 1), OldL, NewL
    M.ReplaceLine L, NewL
Next
End Sub

Sub RplLines(M As CodeModule, Lno&, OldLines$, NewLines$)
DltLines M, Lno, OldLines
M.InsertLines Lno, NewLines
End Sub

Private Sub Z_DltLinzFeis()
Dim A As Feis
'A = MthFeiszMth(Md("Md_"), "XXX")
DltLinzFeis Md("Md_"), A
End Sub

Sub RenTo(FmCmpn, ToNm)
If HasCmpzP(CPj, ToNm) Then Inf CSub, "CmpToNm exist", "ToNm", ToNm: Exit Sub
Cmp(FmCmpn).Name = ToNm
End Sub
Sub Ren(NewCmpn)
CCmp.Name = NewCmpn
End Sub
Function DftPj(P As VBProject) As VBProject
If IsNothing(P) Then
    Set DftPj = CPj
Else
    Set DftPj = P
End If
End Function
Sub RenMdzPfx(FmPfx$, ToPfx$, Optional Pj As VBProject)
Dim P As VBProject: Set P = DftPj(Pj)
Dim C As VBComponent
For Each C In P.VBComponents
    If HasPfx(C.Name, FmPfx) Then
        RenMd C.CodeModule, RplPfx(C.Name, FmPfx, ToPfx)
    End If
Next
End Sub

Sub RenMd(M As CodeModule, NewNm$)
If HasMd(PjzM(M), NewNm) Then
    Debug.Print "New mdn[" & NewNm & "] exist, cannot rename"
    Exit Sub
End If
M.Parent.Name = NewNm
End Sub

Sub MthKeyDrFny()

End Sub

Private Function IfUnRmkMd(M As CodeModule) As Boolean
Debug.Print "UnRmk " & M.Parent.Name,
If Not IsRmkzMd(M) Then
    Debug.Print "No need"
    Exit Function
End If
Debug.Print "<===== is unmarked"
Dim J%, L$
For J = 1 To M.CountOfLines
    L = M.Lines(J, 1)
    If Left(L, 1) <> "'" Then Stop
    M.ReplaceLine J, Mid(L, 2)
Next
IfUnRmkMd = True
End Function

Function IsLinRmk(S) As Boolean
IsLinRmk = FstChr(LTrim(S)) = "'"
End Function

Function IsRmkzS(Src$()) As Boolean
Dim L: For Each L In Itr(Src)
    If Not IsLinRmk(L) Then Exit Function
Next
IsRmkzS = True
End Function

Function IsRmkzMd(M As CodeModule) As Boolean
Dim J%, L$
For J = 1 To M.CountOfLines
    If Left(M.Lines(J, 1), 1) <> "'" Then Exit Function
Next
IsRmkzMd = True
End Function

Sub Rmk()
RmkMd CMd
End Sub

Sub RmkAllMd()
Dim I, Md As CodeModule
Dim NRmk%, Skip%
For Each I In CPj.VBComponents
    If Md.Name <> "LibIdeRmkMd" Then
        If RmkMd(CvMd(I)) Then
            NRmk = NRmk + 1
        Else
            Skip = Skip + 1
        End If
    End If
Next
Debug.Print "NRmk"; NRmk
Debug.Print "SKip"; Skip
End Sub

Private Function RmkMd(M As CodeModule) As Boolean
Debug.Print "Rmk " & M.Parent.Name,
If IsRmkzMd(M) Then
    Debug.Print " No need"
    Exit Function
End If
Debug.Print "<============= is remarked"
Dim J%
For J = 1 To M.CountOfLines
    M.ReplaceLine J, "'" & M.Lines(J, 1)
Next
RmkMd = True
End Function

Sub UnRmk()
IfUnRmkMd CMd
End Sub

Sub UnRmkAllMd()
Dim C As VBComponent
Dim NUnRmk%, Skip%
For Each C In CPj.VBComponents
    If IfUnRmkMd(C.CodeModule) Then
        NUnRmk = NUnRmk + 1
    Else
        Skip = Skip + 1
    End If
Next
Debug.Print "NUnRmk"; NUnRmk
Debug.Print "SKip"; Skip
End Sub

