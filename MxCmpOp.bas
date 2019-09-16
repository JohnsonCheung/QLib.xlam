Attribute VB_Name = "MxCmpOp"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxCmpOp."

Sub AddCls(Clsnn$) 'To CPj
AddCmpzMul CPj, vbext_ct_ClassModule, Clsnn
JmpCmp T1(Clsnn)
End Sub

Sub AddCmpSfx(P As VBProject, Sfx)
If P.Protection = vbext_pp_locked Then Exit Sub
Dim C As VBComponent
For Each C In P.VBComponents
    RenCmp C, C.Name & Sfx
Next
End Sub

Sub AddCmpSfxP(Sfx)
AddCmpSfx CPj, Sfx
End Sub

Sub AddCmpzEmp(P As VBProject, Ty As vbext_ComponentType, NM)
If HasCmpzP(P, NM) Then InfLin CSub, "Cmpn exist in Pj", "Cmpn Pjn", NM, P.Name: Exit Sub
P.VBComponents.Add(Ty).Name = NM ' no CStr will break
End Sub

Sub AddCmpzMul(P As VBProject, T As vbext_ComponentType, Cmpnn$)
Dim N: For Each N In ItrzSS(Cmpnn)
    AddCmpzEmp P, T, N
Next
End Sub

Sub AddCmpzSrc(P As VBProject, NM, SrcL$)
AddCmpzEmp P, vbext_ct_StdModule, NM
ApdLines MdzPN(P, NM), SrcL
End Sub

Sub AddMod(Modnn$)
AddCmpzMul CPj, vbext_ct_StdModule, Modnn
JmpCmp T1(Modnn)
End Sub

Sub AddModzPj(P As VBProject, Modn)
AddCmpzEmp P, vbext_ct_StdModule, Modn
End Sub

Sub ApdLines(M As CodeModule, Lines$)
If Lines = "" Then Exit Sub
M.InsertLines M.CountOfLines + 1, Lines '<=====
End Sub

Sub ApdLineszoInf(M As CodeModule, Lines$)
Dim Bef&, Aft&, Exp&, Cnt&
Bef = M.CountOfLines
ApdLines M, Lines
Aft = M.CountOfLines
Cnt = LinCnt(Lines)
Exp = Bef + Cnt
If Exp <> Aft Then
    Thw CSub, "After copy line count are inconsistents, where [Md], [LinCnt-Bef-Cpy], [LinCnt-of-lines], [Exp-LinCnt-Aft-Cpy], [Act-LinCnt-Aft-Cpy], [Lines]", _
        Mdn(M), Bef, Cnt, Exp, Aft, Lines
End If
End Sub

Sub ApdLy(M As CodeModule, Ly$())
ApdLines M, JnCrLf(Ly)
End Sub

Sub ClrTmpMod()
Dim N
For Each N In TmpModNyzP(CPj)
    If HasPfx(Md(N), "TmpMod") Then RmvCmpzN N
Next
End Sub

Function CpyMd(M As CodeModule, ToM As CodeModule) As Boolean
'Ret : Cpy @M to @ToM and  both must exist @@
CpyMd = RplMd(ToM, SrcL(M))
End Function

Function DftMd(M As CodeModule) As CodeModule
If IsNothing(M) Then
   Set DftMd = CMd
Else
   Set DftMd = M
End If
End Function

Sub DltCmpzPjn(P As VBProject, Mdn)
If Not HasCmpzP(P, Mdn) Then Exit Sub
P.VBComponents.Remove P.VBComponents(Mdn)
End Sub

Sub EnsCls(P As VBProject, Clsn)
EnsCmpzPTN P, vbext_ct_ClassModule, Clsn
End Sub

Sub EnsClsLines(Clsn$, ClsLines$)
EnsCls CPj, Clsn
EnsModLines Md(Clsn), ClsLines
End Sub

Sub EnsCmpzPTN(P As VBProject, Ty As vbext_ComponentType, NM)
If Not HasCmpzP(P, NM) Then AddCmpzEmp P, Ty, NM
End Sub

Sub EnsLines(Md As CodeModule, Mthn, Mthl$)
Dim OldMthL$: OldMthL = MthLzM(Md, Mthn)
If OldMthL = Mthl Then
    Debug.Print FmtQQ("EnsMd: Mth(?) in Md(?) is same", Mthn, Mdn(Md))
End If
RmvMthzMN Md, Mthn
ApdLines Md, Mthl
Debug.Print FmtQQ("EnsMd: Mth(?) in Md(?) is replaced <=========", Mthn, Mdn(Md))
End Sub

Sub EnsMod(P As VBProject, Modn)
EnsCmpzPTN P, vbext_ct_StdModule, Modn
End Sub

Sub EnsModLines(M As CodeModule, Lines$)
If Lines = SrcL(M) Then Inf CSub, "Same module lines, no need to replace", "Mdn", Mdn(M): Exit Sub
RplMd M, Lines
End Sub

Sub EnsModzPN(P As VBProject, Mdn)
EnsCmpzPTN P, vbext_ct_StdModule, Mdn
End Sub

Function HasCmpzN(Cmpn) As Boolean
HasCmpzN = HasCmpzP(CPj, Cmpn)
End Function

Function InsDcl(M As CodeModule, Dcl$) As CodeModule
M.InsertLines FstMthLnozM(M), Dcl
Debug.Print FmtQQ("MdInsDcl: Module(?) a DclLin is inserted", Mdn(M))
End Function

Sub RenCmp(A As VBComponent, NewNm$)
If HasCmpzN(NewNm) Then
    InfLin CSub, "New cmp exists", "OldCmp NewCmp", A.Name, NewNm
Else
    A.Name = NewNm
End If
End Sub

Sub RenCmpOfAddPfx(A As VBComponent, AddPfx$)
A.Name = AddPfx & A.Name
End Sub

Sub RenCmpOfRplPfx(A As VBComponent, FmPfx$, ToPfx$)
If HasPfx(A.Name, FmPfx) Then
    A.Name = RplPfx(A.Name, FmPfx, ToPfx)
End If
End Sub

Sub RmvCmp(A As VBComponent)
A.Collection.Remove A
End Sub

Sub RmvCmpzN(Cmpn)
RmvCmp Cmp(Cmpn)
End Sub

Sub RmvMd(MdDn)
RmvMdzMd Md(MdDn)
End Sub

Sub RmvMdzMd(M As CodeModule)
Dim N$, P$
    N = Mdn(M)
    P = PjnzM(M)
'Debug.Print FmtQQ("RmvMd: Before Md(?) is deleted from Pj(?)", M, P)
M.Parent.Collection.Remove M.Parent
Debug.Print FmtQQ("RmvMd: Md(?) is deleted from Pj(?)", N, P)
End Sub

Sub RmvMdzPfx(Pfx$)
Dim Ny$(): Ny = AwPfx(MdNyP, Pfx)
If Si(Ny) = 0 Then InfLin CSub, "no module begins with " & Pfx: Exit Sub
Brw Ny
Dim N
If Cfm("Rmv those Md as show in the notepad?") Then
    For Each N In Ny
        RmvMd Md(N)
    Next
End If
End Sub

Sub RmvModPfx(Pj As VBProject, Pfx$)
Dim C As VBComponent
For Each C In Pj.VBComponents
    If HasPfx(C.Name, Pfx) Then
        RenCmp C, RmvPfx(C.Name, Pfx)
    End If
Next
End Sub

Sub ClrMd(M As CodeModule)
If M.CountOfLines > 0 Then
    M.DeleteLines 1, M.CountOfLines
End If
End Sub

Private Function SampDiMdnqSrcL() As Dictionary
Set SampDiMdnqSrcL = New Dictionary
Dim C As VBComponent: For Each C In CPj.VBComponents
    SampDiMdnqSrcL.Add C.Name, SrcL(C.CodeModule) & vbCrLf & "'"
Next
End Function

Private Sub Z_RplMdzD()
RplMdzD CPj, SampDiMdnqSrcL
End Sub

Private Function IsStrAtSpcCrLf(S, At) As Boolean
IsStrAtSpcCrLf = IsAscSpcCrLf(AscAt(S, At))
End Function

Private Function AscAt%(S, At)
AscAt = Asc(Mid(S, At, 1))
End Function

Private Function IsAscSpcCrLf(Asc%)
Select Case True
Case Asc = 13, Asc = 10, Asc = 32: IsAscSpcCrLf = True
End Select
End Function

Private Function LineszRTrim$(Lines)
Dim At&
For At = Len(Lines) To 1 Step -1
    If Not IsStrAtSpcCrLf(Lines, At) Then LineszRTrim = Left(Lines, At): Exit Function
Next
End Function

Sub RplMdzD(P As VBProject, DiMdnqSrcL As Dictionary)
'Ret : #Rpl-Md-By-Di-Mdn-SrcL# @@
Dim Mdn: For Each Mdn In DiMdnqSrcL.Keys
    RplMd P.VBComponents(Mdn).CodeModule, DiMdnqSrcL(Mdn)
Next
End Sub

Private Function Si&(A)
On Error Resume Next
Si = UBound(A) + 1
End Function

Private Function LinCnt&(Lines$)
LinCnt = SubStrCnt(Lines, vbLf) + 1
End Function

Private Function SubStrCnt&(S, SubStr$)
Dim P&: P = 1
Dim O&, L%
L = Len(SubStr)
While P > 0
    P = InStr(P, S, SubStr)
    If P = 0 Then SubStrCnt = O: Exit Function
    O = O + 1
    P = P + L
Wend
End Function


Private Sub Z_RplMd()
Dim M As CodeModule: Set M = Md("QDao_Def_NewTd")
RplMd M, SrcL(M) & vbCrLf & "'"
End Sub

Private Function SrcL$(M As CodeModule)
':SrcL: :Lines #Src-Lines#
SrcL = Join(Src(M), vbCrLf) & vbCrLf
End Function

Private Function SrcLzM$(M As CodeModule)
If M.CountOfLines > 0 Then
    SrcLzM = M.Lines(1, M.CountOfLines)
End If
End Function

Private Function Src(M As CodeModule) As String()
Src = SplitCrLf(SrcLzM(M))
End Function

Private Function SplitCrLf(S) As String()
SplitCrLf = Split(Replace(S, vbCr, ""), vbLf)
End Function

Function RTrimMd(M As CodeModule, Optional Upd As EmUpd, Optional Osy) As Boolean

End Function

Function RplMd(M As CodeModule, NewL$) As Boolean
Dim Mdn$: Mdn = M.Parent.Name
Select Case Mdn
Case "QIde_B_CmpOp", "QVb_Dta_VbRpt", "QIde_B_Md", "QIde_Mth_CntMth": Exit Function
End Select

Dim OldL$: OldL = SrcL(M)
Dim IsSam As Boolean: IsSam = LineszRTrim(OldL) = LineszRTrim(NewL)
Dim Msg$
    Dim OldC As String * 4: RSet OldC = LinCnt(OldL)
    Msg = Replace("RplMd: OldCnt(?) ", "?", OldC)
    If IsSam Then
        Msg = Msg & "             " & Mdn & vbTab & "<--- Same"
    Else
        Dim NewC As String * 4: RSet NewC = LinCnt(NewL)
        Msg = Msg & Replace("NewCnt(?) ", "?", NewC) & Mdn
    End If
    Debug.Print Msg
If IsSam Then Exit Function

ClrMd M:
M.InsertLines 1, NewL
RplMd = True
End Function

Private Sub PushI(O, M)
Dim N&
N = Si(O)
ReDim Preserve O(N)
O(N) = M
End Sub

Sub RenModPfx(FmPfx$, ToPfx$)
RenModPfxzP CPj, FmPfx, ToPfx
End Sub

Sub RenModPfxzP(Pj As VBProject, FmPfx$, ToPfx$)
Dim C As VBComponent, N$
For Each C In Pj.VBComponents
    If C.Type = vbext_ct_StdModule Then
        If HasPfx(C.Name, FmPfx) Then
            RenCmp C, RplPfx(C.Name, FmPfx, ToPfx)
        End If
    End If
Next
End Sub

Function SetCmpNm(A As VBComponent, NM, Optional Fun$ = "SetCmpNm") As VBComponent
Dim Pj As VBProject
Set Pj = PjzC(A)
If HasCmpzP(Pj, NM) Then
    Thw Fun, "Cmp already Has", "Cmp Has-in-Pj", NM, Pj.Name
End If
If Pj.Name = NM Then
    Thw Fun, "Cmpn same as Pjn", "Cmpn", NM
End If
A.Name = NM
Set SetCmpNm = A
End Function

Private Sub Z()
MIde__Mth:
End Sub

Sub ChgToCls(FmModn$)
If Not HasCmp(FmModn) Then InfLin CSub, "Mod not exist", "Mod", FmModn: Exit Sub
If Not IsMod(Md(FmModn)) Then InfLin CSub, "It not Mod", "Mod", FmModn: Exit Sub
Dim T$: T = Left(FmModn & "_" & Format(Now, "HHMMDD"), 31)
Md(FmModn).Name = T
AddCls FmModn
Md(FmModn).AddFromString SrcL(Md(T))
RmvCmpzN T
End Sub