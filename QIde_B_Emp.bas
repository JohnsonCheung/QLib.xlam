Attribute VB_Name = "QIde_B_Emp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Md_Emp."
Private Const Asm$ = "QIde"
Function IsMdEmp(M As CodeModule) As Boolean
If M.CountOfLines = 0 Then IsMdEmp = True: Exit Function
Dim J&, L$
For J = 1 To M.CountOfLines
    If Not IsLinEmpSrc(M.Lines(J, 1)) Then Exit Function
Next
IsMdEmp = True
End Function

Sub RmvEmpMd()
Dim N
For Each N In Itr(EmpMdNy)
    RmvCmp CPj.VBComponents(N)
Next
End Sub

Property Get EmpMdNy() As String()
EmpMdNy = EmpMdNyzP(CPj)
End Property

Function EmpMdNyzP(P As VBProject) As String()
Dim C As VBComponent
For Each C In P.VBComponents
    If IsCmpzMd(C) Then
        If IsMdEmp(C.CodeModule) Then
            PushI EmpMdNyzP, C.Name
        End If
    End If
Next
End Function

Private Sub Z_IsMdEmp()
Dim M As CodeModule
'GoSub T1
'GoSub T2
GoSub T3
Exit Sub
T3:
    Debug.Assert IsMdEmp(Md("Dic"))
    Return
T2:
    Set M = Md("Module2")
    Ept = True
    GoTo Tst
T1:
    '
    Dim T$, P As VBProject
        Set P = CPj
        T = TmpNm
    '
'    Set M = PjAdd_Md(P, T)
    Ept = True
    GoSub Tst
    DltCmpzPjn P, T
    Return
Tst:
    Act = IsMdEmp(M)
    C
    Return
End Sub

Function IsEmpSrc(A$()) As Boolean
Dim L
For Each L In Itr(A)
    If Not IsLinEmpSrc(L) Then Exit Function
Next
IsEmpSrc = True
End Function

Function IsLinEmpSrc(A) As Boolean
IsLinEmpSrc = True
If HasPfx(A, "Option ") Then Exit Function
Dim L$: L = Trim(A)
If L = "" Then Exit Function
IsLinEmpSrc = False
End Function

Function EmpMdNyzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy EmpMdNyzV, EmpMdNyzP(P)
Next
End Function


Function IsNoMthMd(M As CodeModule) As Boolean
Dim J&
For J = M.CountOfDeclarationLines + 1 To M.CountOfLines
    If IsLinMth(M.Lines(J, 1)) Then Exit Function
Next
IsNoMthMd = True
End Function

