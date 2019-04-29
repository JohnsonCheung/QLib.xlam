Attribute VB_Name = "MIde_Mth_Dic"
Option Explicit
Public Type DicPair
    Dic1 As Dictionary
    Dic2 As Dictionary
End Type
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

Function MthDicInPj(Optional WiTopRmk As Boolean)

End Function

Function S1S2_Of_TopRmk_And_MthLines(MthLinesWiTopRmk$) As S1S2
Dim N%
    Stop '
With AyabByN(SplitCrLf(MthLinesWiTopRmk), N)
    Dim A$, B$
    A = JnCrLf(.A)
    B = JnCrLf(.B)
    Set S1S2_Of_TopRmk_And_MthLines = S1S2(A, B)
End With
End Function
Function DicPair_Of_MthDicWoTopRmk_And_TopRmkDic(MthDicWiTopRmk As Dictionary) As DicPair
Dim D1 As New Dictionary, D2 As Dictionary, K
    For Each K In MthDicWiTopRmk.Keys
        With S1S2_Of_TopRmk_And_MthLines(MthDicWiTopRmk(K))
            If .S1 <> "" Then
                D2.Add K, .S1
            End If
            D1.Add K, .S2
        End With
    Next
With DicPair_Of_MthDicWoTopRmk_And_TopRmkDic
    Set .Dic1 = D1
    Set .Dic2 = D2
End With
End Function
Function MthDicInPj_WiCache(Optional WiTopRmk As Boolean)
Static X As Boolean, TopRmkDic As New Dictionary, MthDicWoTopRmk As New Dictionary
If Not X Then
    X = True
    With DicPair_Of_MthDicWoTopRmk_And_TopRmkDic(MthDicInPj(WiTopRmk:=True))
        Set MthDicWoTopRmk = .Dic1
        Set TopRmkDic = .Dic2
    End With
End If
If WiTopRmk Then
    Set MthDicInPj = MthDicByDicDic(MthDicWoTopRmk, TopRmkDic)
Else
    Set MthDicInPj = MthDicWoTopRmk
End If
End Function

Function MthDicByDicDic(MthDicWoTopRmk As Dictionary, TopRmkDic As Dictionary) As Dictionary
Dim K, O As New Dictionary
For Each K In MthDicWoTopRmk.Keys
    If TopRmkDic.Exists(K) Then
        O.Add K, TopRmkDic(K) & vbCrLf & MthDicWoTopRmk(K)
    Else
        O.Add K, MthDicWoTopRmk(K)
    End If
Next
Set MthDicByDicDic = O
End Function

Function MthDiczPj(A As VBProject, Optional WiTopRmk As Boolean) As Dictionary
Dim O As New Dictionary, I
For Each I In MdItr(A)
    PushDic O, MthDiczMd(CvMd(I), WiTopRmk)
Next
Set MthDiczPj = O
End Function

Function MthDicInMd() As Dictionary
Set MthDicInMd = MthDiczMd(CurMd)
End Function

Function MthDiczMd(A As CodeModule, Optional WiTopRmk As Boolean) As Dictionary
Set MthDiczMd = AddDicKeyPfx(MthDic(Src(A), WiTopRmk), MdQNmzMd(A) & ".")
End Function

Function MthNmDic(Src$()) As Dictionary 'Key is MthNm.  One PrpNm may have 2 PrpMth: (Get & Set) or (Get & Let)
Dim D As Dictionary: Set D = MthDic(Src): 'Brw LyzNNAp("Src MthDic", Src, FmtDic(D)): Stop
Dim O As New Dictionary, MthDNm$, I
For Each I In D.Keys
    MthDNm = I
    AddDiczApp O, MthNmzMthDNm(MthDNm), D(MthDNm), vbCrLf & vbCrLf
Next
Set MthNmDic = O
End Function

Function MthDic(Src$(), Optional WiTopRmk As Boolean) As Dictionary 'Key is MthDNm, Val is MthLinesWiTopRmk
Dim Ix&, O As New Dictionary, Lines$, DNm$, I
O.Add "*Dcl", Dcl(Src)
For Each I In MthIxItr(Src)
    Ix = I
    DNm = MthDNmzLin(Src(Ix))
    Lines = MthLinesBySrcFm(Src, Ix, WiTopRmk:=WiTopRmk)
    If Lines = "" Then Stop
    O.Add DNm, Lines
Next
Set MthDic = O
End Function
