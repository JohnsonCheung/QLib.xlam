Attribute VB_Name = "MIde_Mth_Rmk"
Option Explicit
Const CMod$ = "MIde_Mth_Rmk."
Function MthCxtFTIxMdMth(A As CodeModule, MthNm) As FTIx

End Function
Sub UnRmkMdMth(A As CodeModule, MthNm$)
UnRmkMdFTIx A, MthCxtFTIxMdMth(A, MthNm)
End Sub
Sub UnRmkMdFTIx(A As CodeModule, B As FTIx)

End Sub
Sub RmkMdFTIx(A As CodeModule, B As FTIx)

End Sub
Sub RmkMdMth(A As CodeModule, MthNm$)
RmkMdFTIx A, MthCxtFTIxMdMth(A, MthNm)
End Sub

Private Sub ZZ_RmkMdMth()
Dim Md As CodeModule, MthNm$
'            Ass LineszVbl(MthLines(M)) = "Property Get ZZA()|End Property||Property Let YYA(V)||End Property"
'RmkMth M:   Ass LineszVbl(MthLines(M)) = "Property Get ZZA()|Stop '|End Property||Property Let YYA(V)|Stop '|'|End Property"
'UnRmkMth M: Ass LineszVbl(MthLines(M)) = "Property Get ZZA()|End Property||Property Let YYA(V)||End Property"
End Sub
Function NxtSrcIx&(Src$(), Ix&)
Const CSub$ = CMod & "NxtSrcIx"
Dim J&
For J = Ix To UB(Src)
    If LasChr(Src(J)) <> "_" Then
        NxtSrcIx = J
        Exit Function
    End If
Next
Thw CSub, "All line From Ix is Src has _ as LasChr", "Ix Src", Ix, AyAddIxPfx(Src, 1)
End Function
Private Function MthCxtFTIxSrcFTIx(Src$(), MthFTIx As FTIx) As FTIx
Set MthCxtFTIxSrcFTIx = FTIx(NxtSrcIx(Src, MthFTIx.FmIx), MthFTIx.ToIx - 1)
End Function

Function NxtMdLno&(A As CodeModule, Lno&)
Const CSub$ = CMod & "NxtMdLno"
Dim J&
For J = Lno To A.CountOfLines
    If LasChr(A.Lines(Lno, 1)) <> "_" Then
        NxtMdLno = J
        Exit Function
    End If
Next
Thw CSub, "All line From Lno has _ as LasChr", "Lno Md Src", Lno, MdNm(A), AyAddIxPfx(Src(A), 1)
End Function
Sub MdFTIx_UnRmk(A As CodeModule, B As FTIx)
If Not Src_IsRmked(LyMdFTIx(A, B)) Then Exit Sub
Dim J%, L$
For J = NxtMdLno(A, B.FmNo) To B.ToNo - 1
    L = A.Lines(J, 1)
    If Left(L, 1) <> "'" Then Stop
    A.ReplaceLine J, Mid(L, 2)
Next
End Sub
Sub MdFTLngSeq_0UnRmk(A As CodeModule, B() As FTIx)
Dim J%
For J = 0 To UB(B)
    MdFTIx_UnRmk A, B(J)
Next
End Sub

Sub MdFTIxAyRmk(A As CodeModule, B() As FTIx)
Dim J%
For J = 0 To UB(B)
    MdFTIx_Rmk A, B(J)
Next
End Sub

Sub MdFTIx_Rmk(A As CodeModule, B As FTIx)
If MdFTIx_IsRmked(A, B) Then Exit Sub
Dim J%
For J = 0 To UB(B)
    A.ReplaceLine J, "'" & A.Lines(J, 1)
Next
End Sub
Function MdFTIx_IsRmked(A As CodeModule, B As FTIx) As Boolean
MdFTIx_IsRmked = Src_IsRmked(LyMdFTIx(A, B))
End Function
Function Src_IsRmked(A$()) As Boolean
If Sz(A) = 0 Then Exit Function
If Not HasPfx(A(0), "Stop '") Then Exit Function
Dim L
For Each L In Itr(A)
    If Left(L, 1) <> "'" Then Exit Function
Next
Src_IsRmked = True
End Function


Function MthFTIxAyCxtFT(Src$(), Mth As FTIx) As FTIx
'Src -> X:MthFmno -> MthCxtFTIxAyNo
With Mth
    Dim Ix%
    For Ix = .FmIx To .ToIx
        If Not LasChr(Src(Ix)) = "_" Then
            Ix = Ix + 1
            Exit For
        End If
    Next
'    Set MthFTIxAyCxtFT = FTIx(Ix, .ToIx - 1)
End With
End Function

Function MthLyCxt(MthLy$()) As String()
'MthLyCxt = CxtFTIx(MthLy, FTIx(1, Sz(MthLy)))
End Function

Function SrcMthNm_CxtFTIxAy(A$(), MthNm$) As FTIx()
Dim P() As FTIx
Dim Ix() As FTIx: 'Ix = SrcMthNmFT(A, MthNm)
'SrcMthNm_CxtFTIxAy = AyMapPX_Into(Ix, "CxtFTIx", A, P)
End Function

Private Sub ZZ_MthCxtFTIxAy _
 _
()

Dim I
For Each I In MthCxtFTIxMdMth(CurMd, CurMthNm)
    With CvFTIx(I)
        Debug.Print .FmNo, .ToNo
    End With
Next
End Sub



