Attribute VB_Name = "MVb_Ay_Map_Align_Pm"
Function AyAlignPm(Ay, PmStr$) As String() 'PmStr [FF..] [AlignNCol:FF..] ..
Dim T1, D As Dictionary
Set D = T1ToAlignNColDic(PmStr)
For Each T1 In D
    PushIAy AyAlignPm, AyAlignPmzT1(Ay, T1, D(T1))
Next
End Function

Private Function AyAlignPmzT1(Ay, T1, AlignNCol) As String()
AyAlignPmzT1 = AyAlignNTerm(AywT1(Ay, T1), CInt(AlignNCol))
End Function

Private Function T1ToAlignNColDic(PmStr$) As Dictionary
Set T1ToAlignNColDic = New Dictionary
    Dim Ay$(), F, D As Dictionary
    Ay = TermAy(PmStr)
    Set D = T1ToAlignNColDiczNoSrt(Ay)
    For Each F In FnyzFF(Ay(0))
        If D.Exists(F) Then
            T1ToAlignNColDic.Add F, D(F)
        Else
            T1ToAlignNColDic.Add F, 1
        End If
    Next
End Function

Private Function T1ToAlignNColDiczNoSrt(PmLy$()) As Dictionary
Dim J%, W%
Set T1ToAlignNColDiczNoSrt = New Dictionary
For J = 2 To UB(PmLy)
    With Brk(PmLy(J), ":")
        W = .S1
        For Each F In FnyzFF(.S2)
            T1ToAlignNColDiczNoSrt.Add F, W
        Next
    End With
Next
End Function

