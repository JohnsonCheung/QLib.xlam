Attribute VB_Name = "QIde_Md_Op_RmkMd"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Md_Op_Rmk."
Private Const Asm$ = "QIde"

Private Function IfUnRmkMd(M As CodeModule) As Boolean
Debug.Print "UnRmk " & M.Parent.Name,
If Not IsRmkedMd(M) Then
    Debug.Print "No need"
    Exit Function
End If
Debug.Print "<===== is unmarked"
Dim J%, L$
For J = 1 To M.CountOfLines
    L = M.Lines(J, 1)
    If Left(L, 1) <> "'" Then Stop
    M.ReplaceLine J, Mid(L, 2)
Next
IfUnRmkMd = True
End Function

Function IsRmkedMd(M As CodeModule) As Boolean
Dim J%, L$
For J = 1 To M.CountOfLines
    If Left(M.Lines(J, 1), 1) <> "'" Then Exit Function
Next
IsRmkedMd = True
End Function

Sub Rmk()
RmkMd CMd
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

Private Function RmkMd(M As CodeModule) As Boolean
Debug.Print "Rmk " & M.Parent.Name,
If IsRmkedMd(M) Then
    Debug.Print " No need"
    Exit Function
End If
Debug.Print "<============= is remarked"
Dim J%
For J = 1 To M.CountOfLines
    M.ReplaceLine J, "'" & M.Lines(J, 1)
Next
RmkMd = True
End Function

Sub UnRmk()
IfUnRmkMd CMd
End Sub

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
