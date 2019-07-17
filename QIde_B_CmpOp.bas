Attribute VB_Name = "QIde_B_CmpOp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Md_Op_Add_Lines."
Private Const Asm$ = "QIde"

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

Sub AddCmpzEmp(P As VBProject, Ty As vbext_ComponentType, Nm)
P.VBComponents.Add(Ty).Name = Nm ' no CStr will break
End Sub

Sub AddCmpzMul(P As VBProject, Ty As vbext_ComponentType, Cmpnn$)
Dim N
For Each N In ItrzSS(Cmpnn)
    AddCmpzEmp P, Ty, N
Next
End Sub

Sub AddCmpzSrc(P As VBProject, Nm, SrcL$)
AddCmpzEmp P, vbext_ct_StdModule, Nm
ApdLines MdzPN(P, Nm), SrcL
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

Function CpyMd(FmM As CodeModule, ToM As CodeModule) As Boolean
'@FmM & @ToM must exist @@
CpyMd = RplMd(ToM, SrcL(FmM))
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

Sub EnsCmpzPTN(P As VBProject, Ty As vbext_ComponentType, Nm)
If Not HasCmpzP(P, Nm) Then AddCmpzEmp P, Ty, Nm
End Sub

Sub EnsLines(Md As CodeModule, Mthn, MthL$)
Dim OldMthL$: OldMthL = MthLzM(Md, Mthn)
If OldMthL = MthL Then
    Debug.Print FmtQQ("EnsMd: Mth(?) in Md(?) is same", Mthn, Mdn(Md))
End If
RmvMthzMN Md, Mthn
ApdLines Md, MthL
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

Function RplMd(M As CodeModule, NewLines$) As Boolean
Dim OldL$: OldL = SrcL(M)
If LineszRTrim(OldL) = LineszRTrim(NewLines) Then Exit Function
ClrMd M
M.InsertLines 1, NewLines
RplMd = True
End Function

Sub RplModPfx(FmPfx$, ToPfx$)
RplModPfxzP CPj, FmPfx, ToPfx
End Sub

Sub RplModPfxzP(Pj As VBProject, FmPfx$, ToPfx$)
Dim C As VBComponent, N$
For Each C In Pj.VBComponents
    If C.Type = vbext_ct_StdModule Then
        If HasPfx(C.Name, FmPfx) Then
            RenCmp C, RplPfx(C.Name, FmPfx, ToPfx)
        End If
    End If
Next
End Sub

Function SetCmpNm(A As VBComponent, Nm, Optional Fun$ = "SetCmpNm") As VBComponent
Dim Pj As VBProject
Set Pj = PjzC(A)
If HasCmpzP(Pj, Nm) Then
    Thw Fun, "Cmp already Has", "Cmp Has-in-Pj", Nm, Pj.Name
End If
If Pj.Name = Nm Then
    Thw Fun, "Cmpn same as Pjn", "Cmpn", Nm
End If
A.Name = Nm
Set SetCmpNm = A
End Function

Sub SrtM()
SrtzM CMd
End Sub

Sub SrtP()
SrtzP CPj
End Sub

Function SrtzM(M As CodeModule) As Boolean
Const C$ = "QIde_Md_Op_RplMd"
If Mdn(M) = C Then Debug.Print "SrtzM: Skipping..."; C
SrtzM = RplMd(M, SSrcLzM(M))
End Function

Sub SrtzP(P As VBProject)
BackupPj
Dim C As VBComponent
For Each C In P.VBComponents
    SrtzM C.CodeModule
Next
End Sub

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

