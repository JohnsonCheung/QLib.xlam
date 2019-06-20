Attribute VB_Name = "QIde_Md_Op_MdOp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Md_Op_Add_Lines."
Private Const Asm$ = "QIde"
Function InsDcl(M As CodeModule, Dcl$) As CodeModule
M.InsertLines FstMthLnozM(M), Dcl
Debug.Print FmtQQ("MdInsDcl: Module(?) a DclLin is inserted", Mdn(M))
End Function

Sub ApdLy(M As CodeModule, Ly$())
ApdLines M, JnCrLf(Ly)
End Sub

Function TmpMod() As CodeModule
Dim T$: T = TmpNm("TmpMod")
AddModzPN CPj, T
Set TmpMod = Md(T)
End Function
Function TmpModNyzP(P As VBProject) As String()
TmpModNyzP = AywPfx(ModNyzP(P), "TmpMod")
End Function

Sub ClrTmpMod()
Dim N
For Each N In TmpModNyzP(CPj)
    If HasPfx(Md(N), "TmpMod") Then RmvCmpzN N
Next
End Sub

Private Sub Z()
MIde__Mth:
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

Function RplMd(M As CodeModule, NewLines$) As Boolean
Dim OldL$: OldL = SrcL(M)
If LineszRTrim(OldL) = LineszRTrim(NewLines) Then Exit Function
ClrMd M
M.InsertLines 1, NewLines
RplMd = True
End Function

Function SrtzM(M As CodeModule) As Boolean
Const C$ = "QIde_Md_Op_RplMd"
If Mdn(M) = C Then Debug.Print "SrtzM: Skipping..."; C
SrtzM = RplMd(M, SSrcLzM(M))
End Function

Sub SrtP()
SrtzP CPj
End Sub

Sub SrtM()
SrtzM CMd
End Sub

Sub SrtzP(P As VBProject)
BackupPj
Dim C As VBComponent
For Each C In P.VBComponents
    SrtzM C.CodeModule
Next
End Sub

