Attribute VB_Name = "MxIm"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxIm."
Function ImAddSfx(Itr, Sfx$) As String()
Dim V: For Each V In Itr
    Push ImAddSfx, V & Sfx
Next
End Function

Function ImAddPfx(Itr, Pfx$) As String()
Dim V: For Each V In Itr
    Push ImAddPfx, Pfx & V
Next
End Function

