Attribute VB_Name = "QIde_Md_Pfx"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Md_Pfx."
Private Const Asm$ = "QIde"
Sub BrwMdPfx()
BrwDic MdPfxDiKqCnt
End Sub

Function MdPfxSyzP(P As VBProject) As String()
MdPfxSyzP = MdPfxSy(MdNyzP(P))
End Function

Function MdPfxDiKqCntzP(P As VBProject) As Dictionary
Set MdPfxDiKqCntzP = DiKqCnt(SrtAyQ(MdPfxSyzP(P)))
End Function

Function MdPfxDiKqCnt() As Dictionary
Set MdPfxDiKqCnt = MdPfxDiKqCntzP(CPj)
End Function

Function MdPfxSy(MdNy$()) As String()
Dim I, N$
For Each I In MdNy
    N = I
    PushI MdPfxSy, MdPfx(N)
Next
End Function
Function MdPfx$(Mdn)
MdPfx = FstCmlzSng(Mdn)
End Function
