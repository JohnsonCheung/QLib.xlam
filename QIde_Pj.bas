Attribute VB_Name = "QIde_Pj"
Option Compare Text
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Pj."

Function CvPj(I) As VBProject
Set CvPj = I
End Function

Function IsPjn(A) As Boolean
IsPjn = HasEle(PjnyV, A)
End Function

Function Pj(Pjn) As VBProject
Set Pj = CVbe.VBProjects(Pjn)
End Function

Sub RmvPj(Pj As VBProject)
Const CSub$ = CMod & "RmvPj"
On Error GoTo X
Dim Pjn$: Pjn = Pj.Name
Pj.Collection.Remove Pj
Exit Sub
X:
Dim E$: E = Err.Description
WarnLin CSub, FmtQQ("Cannot remove Pj[?] Er[?]", Pjn, E)
End Sub
Function StrOfPjfP$()
StrOfPjfP = LineszFt(PjfP)
End Function
Function PjfP$()
PjfP = Pjf(CPj)
End Function
Function PthP$()
PthP = Pjp(CPj)
End Function
Function PjpP$()
PjpP = Pjp(CPj)
End Function
Function Pjp$(P As VBProject)
Pjp = Pth(Pjf(P))
End Function
Function Pjfn$(P As VBProject)
Pjfn = Fn(Pjf(P))
End Function
Function PjfyV() As String()
PjfyV = PjfyzV(CVbe)
End Function

Function PjfyzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushNonBlank PjfyzV, Pjf(P)
Next
End Function
Function PjfnAyV() As String()
PjfnAyV = PjfnAyzV(CVbe)
End Function
Function PjfnAyzV(A As Vbe) As String()
PjfnAyzV = FnAyzFfny(PjfyzV(A))
End Function

Function Pjf$(P As VBProject)
Pjf = PjfzP(P)
End Function

Function PjfzP$(P As VBProject)
On Error GoTo X
PjfzP = P.Filename
Exit Function
X: Debug.Print FmtQQ("Cannot get Pjf for Pj(?). Err[?]", P.Name, Err.Description)
End Function

Function PjFnn$(P As VBProject)
PjFnn = Fnn(Pjf(P))
End Function

Function MdzPN(P As VBProject, Mdn) As CodeModule
Set MdzPN = P.VBComponents(Mdn).CodeModule
End Function

Sub ActPj(P As VBProject)
Set P.Collection.Vbe.ActiveVBProject = P
End Sub

Sub SavPj(P As VBProject)
Const CSub$ = CMod & "SavPj"
If P.Saved Then
    Debug.Print FmtQQ("SavPj: Pj(?) is already saved", P.Name)
    Exit Sub
End If
'Chk Vbe
    Dim Vbe As Vbe
    Set Vbe = P.Collection.Vbe
    If ObjPtr(Vbe.ActiveVBProject) <> ObjPtr(P) Then Stop: Exit Sub
Dim Fnn$
    Fnn = PjFnn(P)
    If Fnn = "" Then
        Thw CSub, "Pj file name is blank.  The pj needs to saved first in order to have a pj file name", "Pj", P.Name
    End If
ActPj P
'Chk BtnOfSav
    Dim B As CommandBarButton: Set B = BtnOfSav(Vbe)
    If B.Caption <> "&Save " & Fnn Then Thw CSub, "Caption is not expected", "Save-Bottun-Caption Expected", B.Caption, "&Save " & Fnn
B.Execute '<===== Save
If P.Saved Then
    Debug.Print FmtQQ("SavPj: Pj(?) is saved <---------------", P.Name)
Else
    Debug.Print FmtQQ("SavPj: Pj(?) cannot be saved for unknown reason <=================================", P.Name)
End If
End Sub

Private Sub Z_SavPj()
SavPj CPj
End Sub


Private Sub ZZ()
Dim A$
Dim B As Variant
Dim C As VBProject
Dim D As Dictionary
Dim E As vbext_ComponentType
ThwIf_BtnOfCompile A
CvPj B
IsPjn B
Pj B
NModzP C
End Sub
Function IsProtectzvInf(P As VBProject) As Boolean
If Not IsProtect(P) Then Exit Function
InfLin CSub, FmtQQ("Skip protected Pj{?)", P.Name)
IsProtectzvInf = True
End Function
Function IsProtect(P As VBProject) As Boolean
IsProtect = P.Protection = vbext_pp_locked
End Function

Sub BrwPjp()
BrwPth PjpP
End Sub

Function FstMd(P As VBProject) As CodeModule
Dim Cmp As VBComponent
For Each Cmp In P.VBComponents
    If IsCmpzMd(CvCmp(Cmp)) Then
        Set FstMd = Cmp.CodeModule
        Exit Function
    End If
Next
End Function

Function FstMod(P As VBProject) As CodeModule
Dim Cmp As VBComponent
For Each Cmp In P.VBComponents
    If IsMod(Cmp) Then
        Set FstMod = Cmp.CodeModule
        Exit Function
    End If
Next
End Function

Function IsFbaPj(P As VBProject) As Boolean
IsFbaPj = IsFba(Pjf(P))
End Function
