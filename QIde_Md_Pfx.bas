Attribute VB_Name = "QIde_Md_Pfx"
Option Explicit
Private Const CMod$ = "MIde_Md_Pfx."
Private Const Asm$ = "QIde"
Sub BrwMdPfx()
BrwDic MdPfxCntDic
End Sub
Function MdPfxSyzPj(A As VBProject) As String()
MdPfxSyzPj = MdPfxSy(MdNyzPj(A))
End Function

Function MdPfxCntDiczPj(A As VBProject) As Dictionary
Set MdPfxCntDiczPj = CntDic(QSrt1(MdPfxSyzPj(A)))
End Function

Function MdPfxCntDic() As Dictionary
Set MdPfxCntDic = MdPfxCntDiczPj(CurPj)
End Function

Function MdPfxSy(MdNy$()) As String()
Dim I, N$
For Each I In MdNy
    N = I
    PushI MdPfxSy, MdPfx(N)
Next
End Function
Function MdPfx$(MdNm$)
MdPfx = FstCmlzWithSng(MdNm)
End Function
