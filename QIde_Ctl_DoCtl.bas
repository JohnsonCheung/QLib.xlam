Attribute VB_Name = "QIde_Ctl_DoCtl"
Option Explicit
Private Const CMod$ = "MIde_Cmd_Action."
Private Const Asm$ = "QIde"
Sub TileH()
BtnOfTileH.Execute
End Sub
Sub TileV()
BtnOfTileV.Execute
End Sub
Sub Compile(Pjn$)
JmpzP Pj(Pjn)
BtnOfCompile.Execute
End Sub
Sub CompilePj(P As VBProject)
JmpzP P
ThwIf_BtnOfCompile P.Name
With BtnOfCompile
    If .Enabled Then
        .Execute
        Debug.Print P.Name, "<--- Compiled"
    Else
        Debug.Print P.Name, "already Compiled"
    End If
End With
BtnOfTileV.Execute
BtnOfSav.Execute
End Sub

Sub CompileVbe(A As Vbe)
DoItrFun A.VBProjects, "PjCompile"
End Sub

Sub ThwIf_BtnOfCompile(NEPjn$)
Dim Act$, Ept$
Act = BtnOfCompile.Caption
Ept = "Compi&le " & NEPjn
If Act <> Ept Then Thw CSub, "Cur BtnOfCompile.Caption <> Compi&le {Pjn}", "Compile-Btn-Caption Pjn Ept-Btn-Caption", Act, NEPjn, Ept
End Sub

Private Sub Z_PjCompile()
CompilePj CPj
End Sub

