Attribute VB_Name = "QIde_Cmp_Op_Add"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Cmp_Op_Add."
Private Const Asm$ = "QIde"

Sub AddCmpzPTN(P As VBProject, Ty As vbext_ComponentType, Nm)
P.VBComponents.Add(Ty).Name = Nm ' no CStr will break
End Sub

Sub AddModzPN(P As VBProject, Modn)
AddCmpzPTN P, vbext_ct_StdModule, Modn
End Sub

Sub AddMod(Modnn$)
AddCmpzPTNn CPj, vbext_ct_StdModule, Modnn
JmpCmp T1(Modnn)
End Sub
Sub AddCmpzPTNn(P As VBProject, Ty As vbext_ComponentType, Cmpnn$)
Dim N
For Each N In ItrzSS(Cmpnn)
    AddCmpzPTN P, Ty, N
Next
End Sub
Sub AddCls(Clsnn$) 'To CPj
AddCmpzPTNn CPj, vbext_ct_ClassModule, Clsnn
JmpCmp T1(Clsnn)
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

Function HasCmpzN(Cmpn) As Boolean
HasCmpzN = HasCmpzP(CPj, Cmpn)
End Function

Sub AddCmpzPNL(P As VBProject, Nm, SrcL$)
AddCmpzPTN P, vbext_ct_StdModule, Nm
ApdLines MdzPN(P, Nm), SrcL
End Sub

Sub RenCmpOfAddPfx(A As VBComponent, AddPfx$)
A.Name = AddPfx & A.Name
End Sub

Sub RenCmpOfRplPfx(A As VBComponent, FmPfx$, ToPfx$)
If HasPfx(A.Name, FmPfx) Then
    A.Name = RplPfx(A.Name, FmPfx, ToPfx)
End If
End Sub

Sub EnsClsLines(Clsn$, ClsLines$)
EnsCls CPj, Clsn
EnsModLines Md(Clsn), ClsLines
End Sub

Sub EnsCls(P As VBProject, Clsn)
EnsCmpzPTN P, vbext_ct_ClassModule, Clsn
End Sub

Sub EnsCmpzPTN(P As VBProject, Ty As vbext_ComponentType, Nm)
If Not HasCmpzP(P, Nm) Then AddCmpzPTN P, Ty, Nm
End Sub

Sub EnsModLines(M As CodeModule, Lines$)
If Lines = SrcL(M) Then Inf CSub, "Same module lines, no need to replace", "Mdn", Mdn(M): Exit Sub
RplMd M, Lines
End Sub

Sub EnsModzPN(P As VBProject, Mdn)
EnsCmpzPTN P, vbext_ct_StdModule, Mdn
End Sub

Sub EnsMod(P As VBProject, Modn)
EnsCmpzPTN P, vbext_ct_StdModule, Modn
End Sub

Private Sub Z()
Dim A$
Dim B As CodeModule
Dim C As VBProject
Dim D As Variant
Dim E As vbext_ComponentType
AddCls A
AddFun A
EnsMod C, A
End Sub


