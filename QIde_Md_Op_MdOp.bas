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


Function RplMth(M As CodeModule, Mthn, NewL$) As Boolean
'Ret True if Rplaced
Dim Lno&: Lno = MthLnozMM(M, Mthn)
If Not HasMthzM(M, Mthn) Then
    RplMth = True
    M.AddFromString NewL '<===
    Exit Function
End If
Dim OldL$: OldL = MthLineszM(M, Mthn)
If OldL = NewL Then Exit Function
RplMth = True
RmvMth M, Mthn '<==
M.InsertLines Lno, NewL '<==
End Function

Private Sub ZZ()
MIde__Mth:
End Sub

Sub EnsLines(Md As CodeModule, Mthn, MthLines$)
Dim OldMthLines$: OldMthLines = MthLineszM(Md, Mthn)
If OldMthLines = MthLines Then
    Debug.Print FmtQQ("EnsMd: Mth(?) in Md(?) is same", Mthn, Mdn(Md))
End If
RmvMthzMN Md, Mthn
ApdLines Md, MthLines
Debug.Print FmtQQ("EnsMd: Mth(?) in Md(?) is replaced <=========", Mthn, Mdn(Md))
End Sub


