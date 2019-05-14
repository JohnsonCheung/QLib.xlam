Attribute VB_Name = "QIde_Md_Op_RmkMd"
Option Explicit
Private Const CMod$ = "MIde_Md_Op_Rmk."
Private Const Asm$ = "QIde"
Sub Rmk()
RmkMd CMd
End Sub
Sub UnRmk()
IfUnRmkMd CMd
End Sub

Sub RmkAllMd()
Dim I, Md As CodeModule
Dim NRmk%, Skip%
For Each I In CPj.VBComponents
    If Md.Name <> "LibIdeRmkMd" Then
        If RmkMd(CvMd(I)) Then
            NRmk = NRmk + 1
        Else
            Skip = Skip + 1
        End If
    End If
Next
Debug.Print "NRmk"; NRmk
Debug.Print "SKip"; Skip
End Sub

Function IsRmkedMd(A As CodeModule) As Boolean
Dim J%, L$
For J = 1 To A.CountOfLines
    If Left(A.Lines(J, 1), 1) <> "'" Then Exit Function
Next
IsRmkedMd = True
End Function


Sub UnRmkAllMd()
Dim C As VBComponent
Dim NUnRmk%, Skip%
For Each C In CPj.VBComponents
    If IfUnRmkMd(C.CodeModule) Then
        NUnRmk = NUnRmk + 1
    Else
        Skip = Skip + 1
    End If
Next
Debug.Print "NUnRmk"; NUnRmk
Debug.Print "SKip"; Skip
End Sub

Private Function RmkMd(A As CodeModule) As Boolean
Debug.Print "Rmk " & A.Parent.Name,
If IsRmkedMd(A) Then
    Debug.Print " No need"
    Exit Function
End If
Debug.Print "<============= is remarked"
Dim J%
For J = 1 To A.CountOfLines
    A.ReplaceLine J, "'" & A.Lines(J, 1)
Next
RmkMd = True
End Function

Private Function IfUnRmkMd(A As CodeModule) As Boolean
Debug.Print "UnRmk " & A.Parent.Name,
If Not IsRmkedMd(A) Then
    Debug.Print "No need"
    Exit Function
End If
Debug.Print "<===== is unmarked"
Dim J%, L$
For J = 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If Left(L, 1) <> "'" Then Stop
    A.ReplaceLine J, Mid(L, 2)
Next
IfUnRmkMd = True
End Function
