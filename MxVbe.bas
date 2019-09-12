Attribute VB_Name = "MxVbe"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxVbe."
Enum EmSrtLisMd
    EiByMdn
    EiByMdnDes
    EiByNLines
    EiByNLinesDes
End Enum
Function CvVbe(A) As Vbe
Set CvVbe = A
End Function

Sub DmpIsPjSav()
DmpDrs DoIsPjSav(CVbe)
End Sub

Private Function DyoIsPjSav(A As Vbe) As Variant()
Dim I As VBProject
For Each I In A.VBProjects
    PushI DyoIsPjSav, Array(I.Saved, I.Name, I.GenFileName)
Next
End Function

Function DoIsPjSav(A As Vbe) As Drs
DoIsPjSav = DrszFF("IsSav Pjn GenFfn", DyoIsPjSav(A))
End Function

Function PjzV(A As Vbe, Pjn$) As VBProject
Set PjzV = A.VBProjects(Pjn)
End Function

Function PjzPjf(Vbe As Vbe, Pjf) As VBProject
Dim I As VBProject
For Each I In Vbe.VBProjects
    If PjfzP(I) = Pjf Then Set PjzPjf = I: Exit Function
Next
End Function

Sub BrwMd(Optional MdPatn$, Optional SrtBy As EmSrtLisMd)
LisMd MdPatn, SrtBy, OupTy:=EiOtBrw, Top:=0
End Sub

Sub VcMd(Optional MdPatn$, Optional SrtBy As EmSrtLisMd)
LisMd MdPatn, SrtBy, OupTy:=EiOtVc, Top:=0
End Sub

Sub LisMd(Optional MdPatn$, Optional SrtBy As EmSrtLisMd, Optional OupTy As EmOupTy = EmOupTy.EiOtDmp, Optional Top% = 50)
Dim Srt$
Select Case True
Case SrtBy = EiByMdn:       Srt = "Mdn"
Case SrtBy = EiByMdnDes:    Srt = "-Mdn"
Case SrtBy = EiByNLines:    Srt = "Mdn"
Case SrtBy = EiByNLinesDes: Srt = "-NLin"
Case Else:                  Srt = "Mdn"
End Select
Brw FmtCellDrs(SrtDrs(DoMdP, Srt), , Fmt:=EiSSFmt), OupTy:=OupTy
End Sub

Sub SavVbe(A As Vbe)
Dim P As VBProject
For Each P In A.VBProjects
    SavPj P
Next
End Sub

Function PjfyV() As String()
PjfyV = PjfyzV(CVbe)
End Function

Function PjfyzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushNB PjfyzV, Pjf(P)
Next
End Function

Function PjNyV() As String()
PjNyV = PjNyzV(CVbe)
End Function

Function PjNyzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushI PjNyzV, P.Name
Next
End Function

Function SrtRptV() As String()
SrtRptV = SrtRptzV(CVbe)
End Function

Function HasBarzV(A As Vbe, BarNm) As Boolean
HasBarzV = HasItn(A.CommandBars, BarNm)
End Function

Function HasPj(A As Vbe, Pjn$) As Boolean
HasPj = HasItn(A.VBProjects, Pjn)
End Function

Function HasPjfzV(A As Vbe, Pjf) As Boolean
Dim P As VBProject
For Each P In A.VBProjects
    If PjfzP(P) = Pjf Then HasPjfzV = True: Exit Function
Next
End Function

Function SrtRptzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy SrtRptzV, SrtRptzP(P)
Next
End Function

Private Sub Z_VbeFunPfx()
'D Vbe_MthPfx(CVbe)
End Sub

Private Sub Z_MthNyzV()
Brw MthNyzV(CVbe)
End Sub


Private Sub Z()
Dim A
Dim B As Vbe
Dim C$
Dim D As Boolean
Dim XX
CvVbe A
End Sub
