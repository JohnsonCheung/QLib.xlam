Attribute VB_Name = "MIde_Pj"
Option Explicit
Const CMod$ = "MIde_Pj."

Sub ThwIfCompileBtn(NEPjNm$)
Dim Act$, Ept$
Act = CompileBtn.Caption
Ept = "Compi&le " & NEPjNm
If Act <> Ept Then Thw CSub, "Cur CompileBtn.Caption <> Compi&le {PjNm}", "Compile-Btn-Caption PjNm Ept-Btn-Caption", Act, NEPjNm, Ept
End Sub
Function CvPj(I) As VBProject
Set CvPj = I
End Function

Function IsPjNm(A) As Boolean
IsPjNm = HasEle(PjNy, A)
End Function

Function Pj(PjNm) As VBProject
Set Pj = CurVbe.VBProjects(PjNm)
End Function

Sub RmvPj(Pj As VBProject)
Const CSub$ = CMod & "RmvPj"
On Error GoTo X
Dim PjNm$: PjNm = Pj.Name
Pj.Collection.Remove Pj
Exit Sub
X:
Dim E$: E = Err.Description
WarnLin CSub, FmtQQ("Cannot remove Pj[?] Er[?]", PjNm, E)
End Sub
Function StrzCurPjf$()
StrzCurPjf = FtLines(PjfPj)
End Function
Function PjfPj$()
PjfPj = Pjf(CurPj)
End Function
Function PjPth$(A As VBProject)
PjPth = Pth(Pjf(A))
End Function

Function Pjf$(A As VBProject)
On Error GoTo X
Pjf = A.Filename
Exit Function
X: Debug.Print "Cannot get Pjf, Err:"; Err.Description
End Function

Function PjFnn$(A As VBProject)
PjFnn = Fnn(Pjf(A))
End Function

Function IsUsrLibPj(A As VBProject) As Boolean
IsUsrLibPj = IsFxa(A)
End Function

Function MdzPj(A As VBProject, Nm) As CodeModule
Set MdzPj = A.VBComponents(Nm).CodeModule
End Function


Sub Compile()
CompilePj CurPj
End Sub

Sub CompilePj(A As VBProject)
JmpPj A
ThwIfCompileBtn A.Name
With CompileBtn
    If .Enabled Then
        .Execute
        Debug.Print A.Name, "<--- Compiled"
    Else
        Debug.Print A.Name, "already Compiled"
    End If
End With
TileVBtn.Execute
SavBtn.Execute
End Sub

Sub ActPj(A As VBProject)
Set A.Collection.Vbe.ActiveVBProject = A
End Sub

Sub SavPj(A As VBProject)
Const CSub$ = CMod & "SavPj"
If A.Saved Then
    Debug.Print FmtQQ("SavPj: Pj(?) is already saved", A.Name)
    Exit Sub
End If
'Chk Vbe
    Dim Vbe As Vbe
    Set Vbe = A.Collection.Vbe
    If ObjPtr(Vbe.ActiveVBProject) <> ObjPtr(A) Then Stop: Exit Sub
Dim Fnn$
    Fnn = PjFnn(A)
    If Fnn = "" Then
        Thw CSub, "Pj file name is blank.  The pj needs to saved first in order to have a pj file name", "Pj", A.Name
    End If
ActPj A
'Chk SavBtn
    Dim B As CommandBarButton: Set B = SavBtn(Vbe)
    If B.Caption <> "&Save " & Fnn Then Thw CSub, "Caption is not expected", "Save-Bottun-Caption Expected", B.Caption, "&Save " & Fnn
B.Execute '<===== Save
If A.Saved Then
    Debug.Print FmtQQ("SavPj: Pj(?) is saved <---------------", A.Name)
Else
    Debug.Print FmtQQ("SavPj: Pj(?) cannot be saved for unknown reason <=================================", A.Name)
End If
End Sub

Private Sub ZZ_SavPj()
SavPj CurPj
End Sub

Private Sub Z_PjCompile()
CompilePj CurPj
End Sub

Private Sub ZZ()
Dim A$
Dim B As Variant
Dim C As VBProject
Dim D As Dictionary
Dim E As vbext_ComponentType
ThwIfCompileBtn A
CvPj B
IsPjNm B
Pj B
IsUsrLibPj C
MdzPj C, B
NModzPj C
End Sub
Function IsProtectzInfo(A As VBProject) As Boolean
If Not IsProtect(A) Then Exit Function
InfoLin CSub, FmtQQ("Skip protected Pj{?)", A.Name)
IsProtectzInfo = True
End Function
Function IsProtect(A As VBProject) As Boolean
IsProtect = A.Protection = vbext_pp_locked
End Function

Sub BrwPthPj()
BrwPth PthPj
End Sub

Function PjzXls(A As Excel.Application, Fxa) As VBProject
Dim P As VBProject
For Each P In A.Application.Vbe.VBProjects
    If Pjf(P) = Fxa Then
        Set PjzXls = P
        Exit Function
    End If
Next
End Function

Function FstMd(A As VBProject) As CodeModule
Dim Cmp As VBComponent
For Each Cmp In A.VBComponents
    If IsMd(CvCmp(Cmp)) Then
        Set FstMd = Cmp.CodeModule
        Exit Function
    End If
Next
End Function

Function FstMod(A As VBProject) As CodeModule
Dim Cmp As VBComponent
For Each Cmp In A.VBComponents
    If IsMod(Cmp) Then
        Set FstMod = Cmp.CodeModule
        Exit Function
    End If
Next
End Function

Function IsFbaPj(A As VBProject) As Boolean
IsFbaPj = IsFba(Pjf(A))
End Function
