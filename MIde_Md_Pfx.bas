Attribute VB_Name = "MIde_Md_Pfx"
Option Explicit
Sub BrwMdPfx()
BrwDic MdPfxCntDic
End Sub
Function MdPfxAyzPj(A As VBProject) As String()
MdPfxAyzPj = MdPfxAy(MdNyzPj(A))
End Function

Function MdPfxCntDiczPj(A As VBProject) As Dictionary
Set MdPfxCntDiczPj = CntDic(AyQSrt(MdPfxAyzPj(A)))
End Function

Function MdPfxCntDic() As Dictionary
Set MdPfxCntDic = MdPfxCntDiczPj(CurPj)
End Function

Function MdPfxAy(MdNy$()) As String()
Dim I, N$
For Each I In MdNy
    N = I
    PushI MdPfxAy, MdPfx(N)
Next
End Function
Function MdPfx$(MdNm$)
MdPfx = FstCmlzWithSng(MdNm)
End Function
