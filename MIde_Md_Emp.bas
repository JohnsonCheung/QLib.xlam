Attribute VB_Name = "MIde_Md_Emp"
Option Explicit
Function IsEmpMd(A As CodeModule) As Boolean
If A.CountOfLines = 0 Then IsEmpMd = True: Exit Function
Dim J&, L$
For J = 1 To A.CountOfLines
    If Not IsEmpSrcLin(A.Lines(J, 1)) Then Exit Function
Next
IsEmpMd = True
End Function

Sub RmvEmpMd()
Dim N
For Each N In Itr(EmpMdNy)
    RmvCmp CurPj.VBComponents(N)
Next
End Sub

Property Get EmpMdNy() As String()
EmpMdNy = EmpMdNyzPj(CurPj)
End Property

Function EmpMdNyzPj(A As VBProject) As String()
Dim C As VBComponent
For Each C In A.VBComponents
    If IsMd(C) Then
        If IsEmpMd(C.CodeModule) Then
            PushI EmpMdNyzPj, C.Name
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
        Set P = CurPj
        T = TmpNm
    '---
'    Set M = PjAdd_Md(P, T)
    Ept = True
    GoSub Tst
    DltCmpz P, T
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

Function EmpMdNyzVbe(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy EmpMdNyzVbe, EmpMdNyzPj(P)
Next
End Function


Function IsNoMthMd(A As CodeModule) As Boolean
Dim J&
For J = A.CountOfDeclarationLines + 1 To A.CountOfLines
    If IsMthLin(A.Lines(J, 1)) Then Exit Function
Next
IsNoMthMd = True
End Function

