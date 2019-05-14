Attribute VB_Name = "QIde_Md_Pfx"
Option Explicit
Private Const CMod$ = "MIde_Md_Pfx."
Private Const Asm$ = "QIde"
Sub BrwMdPfx()
BrwDic MdPfxCntDic
End Sub

Function MdPfxSyzP(P As VBProject) As String()
MdPfxSyzP = MdPfxSy(MdNyzP(P))
End Function

Function MdPfxCntDiczP(P As VBProject) As Dictionary
Set MdPfxCntDiczP = CntDic(QSrt1(MdPfxSyzP(P)))
End Function

Function MdPfxCntDic() As Dictionary
Set MdPfxCntDic = MdPfxCntDiczP(CPj)
End Function

Function MdPfxSy(MdNy$()) As String()
Dim I, N$
For Each I In MdNy
    N = I
    PushI MdPfxSy, MdPfx(N)
Next
End Function
Function MdPfx$(Mdn)
'MdPfx = FstCmlzWithSng(Mdn)
End Function
