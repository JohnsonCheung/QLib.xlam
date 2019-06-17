Attribute VB_Name = "QIde_Md_Op_RplMd"
Option Explicit
Option Compare Text
Function RplMd(M As CodeModule, NewLines$) As Boolean
Dim OldL$: OldL = SrcLines(M)
If RTrimLines(OldL) = RTrimLines(NewLines) Then Exit Function
ClrMd M
M.InsertLines 1, NewLines
RplMd = True
End Function

Sub SrtzM(M As CodeModule)
Const C$ = "QIde_Md_Op_RplMd"
If Mdn(M) = C Then Debug.Print "SrtzM: Skipping..."; C
RplMd M, SSrcLineszM(M)
End Sub

Sub SrtP()
SrtzP CPj
End Sub

Sub SrtM()
SrtzM CMd
End Sub

Sub SrtzP(P As VBProject)
BackupPj
Dim C As VBComponent
For Each C In P.VBComponents
    SrtzM C.CodeModule
Next
End Sub
