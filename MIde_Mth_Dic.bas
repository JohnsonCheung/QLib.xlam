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


Private Sub Z_MthDiczMd()
BrwDic MthDiczMd(CurMd)
End Sub
Private Sub Z_PjMthDic()
Dim A As Dictionary, V, K
Set A = Pj_MthDic(CurPj)
Ass IsDiczSy(A) '
For Each K In A
    If InStr(K, ".") > 0 Then Stop
    If Si(A(K)) = 0 Then Stop
Next
End Sub

Private Sub Z_PjMthDic1()
Dim A As Dictionary, V, K
Set A = Pj_MthDic(CurPj)
Ass IsDiczSy(A) '
For Each K In A
    If InStr(K, ".") > 0 Then Stop
    If Si(A(K)) = 0 Then Stop
Next
End Sub
Private Sub Z()
Z_PjMthDic
Z_PjMthDic1
MIde_Mth_Dic:
End Sub

Private Sub Z_MthDic()
Dim Src$()
ZZ:
    Src = Sy("Sub AA(): XX : End Sub", "Sub BB()", "XX", "End Sub")
    Brw MthDic(Src)
    Return
End Sub
Function MthDicPj()
Set MthDicPj = MthDiczPj(CurPj)
End Function

Function MthDiczPj(A As VBProject, Optional WiTopRmk As Boolean) As Dictionary
Dim O As New Dictionary, I
For Each I In MdItr(A)
    PushDic O, MthDiczMd(CvMd(I), WiTopRmk)
Next
Set MthDiczPj = O
End Function

Function MthDicMd() As Dictionary
Set MthDicMd = MthDiczMd(CurMd)
End Function

Function MthDiczMd(A As CodeModule, Optional WiTopRmk As Boolean) As Dictionary
Set MthDiczMd = AddDicKeyPfx(MthDic(Src(A), WiTopRmk), MdQNmzMd(A) & ".")
End Function

Function MthNmDic(Src$()) As Dictionary 'Key is MthNm.  One PrpNm may have 2 PrpMth: (Get & Set) or (Get & Let)
Dim D As Dictionary: Set D = MthDic(Src): 'Brw LyzNNAp("Src MthDic", Src, FmtDic(D)): Stop
Dim O As New Dictionary, MthDNm
For Each MthDNm In D.Keys
    AddDiczApp O, MthNmzMthDNm(MthDNm), D(MthDNm), vbCrLf & vbCrLf
Next
Set MthNmDic = O
End Function

Function MthDic(Src$(), Optional WiTopRmk As Boolean) As Dictionary 'Key is MthDNm, Val is MthLinesWiTopRmk
Dim Ix, O As New Dictionary, Lines$, DNm$
O.Add "*Dcl", Dcl(Src)
For Each Ix In MthIxItr(Src)
    DNm = MthDNmzLin(Src(Ix))
    Lines = MthLinesBySrcFm(Src, Ix, WiTopRmk:=WiTopRmk)
    If Lines = "" Then Stop
    O.Add DNm, Lines
Next
Set MthDic = O
End Function
