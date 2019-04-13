Attribute VB_Name = "MIde_Cmp_Op_Add"
Option Explicit
Function AddCmpM(Nm) As VBComponent
Set AddCmpM = AddCmp(Nm, vbext_ct_StdModule)
End Function
Function AddCmpC(Nm) As VBComponent
Set AddCmpC = AddCmp(Nm, vbext_ct_ClassModule)
End Function
Function AddCmp(Nm, Ty As vbext_ComponentType) As VBComponent
Set AddCmp = AddCmpzPj(CurPj, Nm, Ty)
End Function
Function AddCmpzPj(A As VBProject, Nm, Ty As vbext_ComponentType) As VBComponent
If HasCmp(Nm) Then InfLin CSub, FmtQQ("?[?] already exist", ShtCmpTy(Ty), Nm): Exit Function
Dim O As VBComponent
Set O = A.VBComponents.Add(Ty)
O.Name = CStr(Nm) ' no CStr will break
Set AddCmpzPj = O
End Function

Function AddModzPj(A As VBProject, ModNm) As CodeModule
AddCmpzPj A, ModNm, vbext_ct_StdModule
End Function

Sub AddMod(ModNN$)
Dim Sy$(), ModNm
Sy = SySsl(ModNN)
For Each ModNm In Sy
    AddModzPj CurPj, ModNm
Next
JmpCmp Sy(0)
End Sub

Function IsErDmp(Er$()) As Boolean
If Si(Er) = 0 Then Exit Function
D Er
IsErDmp = True
End Function
Sub AddCls(ClsNN$) 'To CurPj
Dim ClsNm, Sy$()
Sy = SySsl(ClsNN)
For Each ClsNm In Sy
    AddCmp ClsNm, vbext_ComponentType.vbext_ct_ClassModule
Next
JmpCmp Sy(0)
End Sub

Sub ApdLines(A As CodeModule, Lines)
If Lines = "" Then Exit Sub
Dim Bef&, Aft&, Exp&, Cnt&
Bef = A.CountOfLines
A.InsertLines A.CountOfLines + 1, Lines '<=====
Aft = A.CountOfLines
Cnt = LinCnt(Lines)
Exp = Bef + Cnt
If Exp <> Aft Then
'    Thw CSub, "After copy line count are inconsistents, where [Md], [LinCnt-Bef-Cpy], [LinCnt-of-lines], [Exp-LinCnt-Aft-Cpy], [Act-LinCnt-Aft-Cpy], [Lines]", _
        MdNm(A), Bef, Cnt, Exp, Aft, Lines
End If
End Sub

Sub AddFun(FunNm$)
ApdLines CurMd, EmpFunLines(FunNm)
JmpMth FunNm
End Sub
Function CmpNew(Nm$, Ty As vbext_ComponentType) As VBComponent
Set CmpNew = CurPj.VBComponents.Add(Ty)
End Function

Function EmpFunLines$(FunNm)
EmpFunLines = FmtQQ("Function ?()|End Function", FunNm)
End Function

Function EmpSubLines$(SubNm)
EmpSubLines = FmtQQ("Sub ?()|End Sub", SubNm)
End Function
Sub AddSub(SubNm$)
ApdLines CurMd, EmpSubLines(SubNm)
JmpMth SubNm
End Sub

Function AddOptExpLinMd(A As CodeModule) As CodeModule
A.InsertLines 1, "Option Explicit"
Set AddOptExpLinMd = A
End Function

Function HasCmp(CmpNm) As Boolean
HasCmp = HasCmpzPj(CurPj, CmpNm)
End Function
Function AddCmpzLines(A As VBProject, Nm, SrcLines$) As VBComponent
Dim O As VBComponent
Set O = AddCmpzPj(A, Nm, vbext_ct_StdModule): If IsNothing(O) Then Stop
ApdLines O.CodeModule, SrcLines
Set AddCmpzLines = O
End Function
Sub RenAddCmpPfx_CmpPfx(A As VBComponent, AddPfx$)
A.Name = AddPfx & A.Name
End Sub
Function ModCmpItr(Pj As VBProject)

End Function

Function ModCmpAy(Pj As VBProject) As VBComponent()
ModCmpAy = IntozItrPEv(ModCmpAy, Pj.VBComponents, "Type", vbext_ct_StdModule)
End Function

Sub RenCmpRplPfx(A As VBComponent, FmPfx$, ToPfx$)
If HasPfx(A.Name, FmPfx) Then
    A.Name = RplPfx(A.Name, FmPfx, ToPfx)
End If
End Sub
Sub CrtPjNmzCmpTy(A As VBProject, Nm, Ty As vbext_ComponentType)
Dim O As VBComponent
Set O = A.VBComponents.Add(Ty)
O.Name = Nm
End Sub
Sub CrtCls(A As VBProject, Nm$)
CrtPjNmzCmpTy A, Nm, vbext_ct_ClassModule
End Sub
Function EnsCls(A As VBProject, ClsNm$) As CodeModule
Set EnsCls = EnsCmp(A, ClsNm, vbext_ct_ClassModule)
End Function

Function EnsCmp(A As VBProject, Nm, Optional Ty As vbext_ComponentType = vbext_ct_StdModule) As VBComponent
If Not HasCmpzPj(A, Nm) Then
    A.VBComponents.Add(Ty).Name = Nm
End If
Set EnsCmp = A.VBComponents(Nm)
End Function

Function EnsMdzPj(A As VBProject, MdNm) As CodeModule
Set EnsMdzPj = EnsCmp(A, MdNm, vbext_ct_StdModule).CodeModule
End Function

Function EnsMod(A As VBProject, ModNm$) As CodeModule
Set EnsMod = EnsCmp(A, ModNm, vbext_ct_StdModule)
End Function

Private Sub ZZ()
Dim A$
Dim B As CodeModule
Dim C As VBProject
Dim D As Variant
Dim E As vbext_ComponentType
Dim F As WhMd
AddCls A
AddFun A
EnsCmp C, D, E
EnsMdzPj C, D
EnsMod C, A
End Sub

Private Sub Z()
End Sub





