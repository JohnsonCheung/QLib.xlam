Attribute VB_Name = "MxBrwMd"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxBrwMd."
Enum EmSrtLisMd
    EiByMdn
    EiByMdnDes
    EiByNLines
    EiByNLinesDes
End Enum

Sub BrwMdP()
BrwDrs DoMdP
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
Brw FmtCellDrs(SrtDrs(DwPatn(DoMdP, "Mdn", MdPatn), Srt), , Fmt:=EiSSFmt), OupTy:=OupTy
End Sub

Sub BrwMd(Optional MdPatn$)
BrwDrs SrtDrs(DwPatn(DoMdP, "Mdn", MdPatn))
End Sub

Sub VcMd(Optional MdPatn$, Optional SrtBy As EmSrtLisMd)
LisMd MdPatn, SrtBy, OupTy:=EiOtVc, Top:=0
End Sub
