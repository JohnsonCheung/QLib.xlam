Attribute VB_Name = "QIde_Vbe"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Vbe."
Private Const Asm$ = "QIde"
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

Sub LisMd(Optional MdPatn$ = ".+", Optional SrtBy As EmSrtLisMd, Optional Oup As EmOupTy)
Dim Srt$
Select Case True
Case SrtBy = EiByMdn: Srt = "Mdn"
Case SrtBy = EiByMdnDes: Srt = "-Mdn"
Case SrtBy = EiByNLines: Srt = "Mdn"
Case SrtBy = EiByNLinesDes: Srt = "-NLines"
Case Else:    Srt = "Mdn"
End Select
Dmp FmtDrs(SrtDrs(DoMdV(MdPatn), Srt), , Fmt:=EiSSFmt), Oup
End Sub

Function DoMdV(Optional MdPatn$ = ".*") As Drs
DoMdV = Drs(FoMd, DyoMd(CVbe, MdPatn))
End Function

Function DoMd(A As Vbe) As Drs
DoMd = Drs(FoMd, DyoMd(A))
End Function

Function FoMd() As String()
FoMd = SyzSS("Mdn NLines")
End Function

Function DyoMd(A As Vbe, Optional MdPatn$ = ".*") As Variant()
Dim Re As RegExp: Set Re = RegExp(MdPatn)
Dim P As VBProject: For Each P In A.VBProjects
    Dim C As VBComponent: For Each C In P.VBComponents
        If Re.Test(C.Name) Then
            PushI DyoMd, DroMd(C.CodeModule)
        End If
    Next
Next
End Function

Function DroMd(M As CodeModule) As Variant()
DroMd = Array(M.Parent.Name, M.CountOfLines)
End Function

Sub SavVbe(A As Vbe)
Dim P As VBProject
For Each P In A.VBProjects
    SavPj P
Next
End Sub

Property Get PjfyV() As String()
PjfyV = PjfyzV(CVbe)
End Property

Function PjfyzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushNB PjfyzV, Pjf(P)
Next
End Function

Function PjnyV() As String()
PjnyV = PjnyzV(CVbe)
End Function

Function PjnyzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushI PjnyzV, P.Name
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

