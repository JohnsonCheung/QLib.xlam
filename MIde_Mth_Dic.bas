Attribute VB_Name = "MIde_Mth_Dic"
Option Explicit
Private Sub ZZ_Pj_MthDic()
BrwDic Pj_MthDic(CurPj)
End Sub

Function Pj_MthDic(A As VBProject) As Dictionary
Dim C As VBComponent
Set Pj_MthDic = New Dictionary
For Each C In A.VBComponents
    PushDic Pj_MthDic, MthDiczMd(C.CodeModule)
Next
End Function

Private Sub ZZ_MdMthDic()
BrwDic MthDiczMd(CurMd)
End Sub


Private Sub Z_MdMthDic()
BrwDic MthDiczMd(CurMd)
End Sub
Private Sub Z_PjMthDic()
Dim A As Dictionary, V, K
Set A = Pj_MthDic(CurPj)
Ass IsDiczSy(A) '
For Each K In A
    If InStr(K, ".") > 0 Then Stop
    If Sz(A(K)) = 0 Then Stop
Next
End Sub

Private Sub Z_PjMthDic1()
Dim A As Dictionary, V, K
Set A = Pj_MthDic(CurPj)
Ass IsDiczSy(A) '
For Each K In A
    If InStr(K, ".") > 0 Then Stop
    If Sz(A(K)) = 0 Then Stop
Next
End Sub
Private Sub Z()
Z_MdMthDic
Z_PjMthDic
Z_PjMthDic1
MIde_Mth_Dic:
End Sub

Private Sub Z_MthDic()
BrwDic MthDic(Src(Md("AAAMod")))
End Sub
Function MthDicPj()
Set MthDicPj = MthDiczPj(CurPj)
End Function
Function MthDiczPj(A As VBProject) As Dictionary
Dim O As New Dictionary, I
For Each I In MdItr(A)
    PushDic O, MthDiczMd(CvMd(I))
Next
Set MthDiczPj = O
End Function
Function MthDicMd() As Dictionary
Set MthDicMd = MthDiczMd(CurMd)
End Function
Function MthDiczMd(A As CodeModule) As Dictionary
Set MthDiczMd = AddDicKeyPfx(MthDic(Src(A)), MdQNmzMd(A) & ".")
End Function

Function MthDic(Src$()) As Dictionary
Dim Ix, O As New Dictionary
O.Add "*Dcl", DclLines(Src)
For Each Ix In MthIxItr(Src)
    O.Add MthDNmzLin(Src(Ix)), MthLineszSrcFm(Src, Ix, WithTopRmk:=True)
Next
Set MthDic = O
End Function
