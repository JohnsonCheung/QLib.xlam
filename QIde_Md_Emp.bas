Attribute VB_Name = "QIde_Md_Emp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Md_Emp."
Private Const Asm$ = "QIde"
Function IsEmpMd(M As CodeModule) As Boolean
If M.CountOfLines = 0 Then IsEmpMd = True: Exit Function
Dim J&, L$
For J = 1 To M.CountOfLines
    If Not IsEmpSrcLin(M.Lines(J, 1)) Then Exit Function
Next
IsEmpMd = True
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
        If IsEmpMd(C.CodeModule) Then
            PushI EmpMdNyzP, C.Name
        End If
    End If
Next
End Function

Private Sub Z_IsEmpMd()
Dim M As CodeModule
'GoSub T1
'GoSub T2
GoSub T3
Exit Sub
T3:
    Debug.Assert IsEmpMd(Md("Dic"))
    Return
T2:
    Set M = Md("Module2")
    Ept = True
    GoTo Tst
T1:
    '-----
    Dim T$, P As VBProject
        Set P = CPj
        T = TmpNm
    '---
'    Set M = PjAdd_Md(P, T)
    Ept = True
    GoSub Tst
    DltCmpzPjn P, T
    Return
Tst:
    Act = IsEmpMd(M)
    C
    Return
End Sub

Function IsEmpSrc(A$()) As Boolean
Dim L
For Each L In Itr(A)
    If Not IsEmpSrcLin(L) Then Exit Function
Next
IsEmpSrc = True
End Function

Function IsEmpSrcLin(A) As Boolean
IsEmpSrcLin = True
If HasPfx(A, "Option ") Then Exit Function
Dim L$: L = Trim(A)
If L = "" Then Exit Function
IsEmpSrcLin = False
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
    If IsMthLin(M.Lines(J, 1)) Then Exit Function
Next
IsNoMthMd = True
End Function

