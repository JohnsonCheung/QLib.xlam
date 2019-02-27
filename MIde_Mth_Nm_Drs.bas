Attribute VB_Name = "MIde_Mth_Nm_Drs"
Option Explicit

Function MthNmDRsVbe(Optional WhStr$) As DRs
Set MthNmDRsVbe = MthNmDrszVbe(CurVbe, WhStr)
End Function

Function MthNmDRsPj(Optional WhStr$) As DRs
Set MthNmDRsPj = MthNmDrszPj(CurPj, WhStr)
End Function

Function MthNmDRsMd(Optional WhStr$) As DRs
Set MthNmDRsMd = MthNmDrszMd(CurMd, WhStr)
End Function

Private Function MthNmDrszMd(M As CodeModule, Optional WhStr$) As DRs
Set MthNmDrszMd = DRs(MthNmFny, MthNmDryzMd(M, WhMthzStr(WhStr)))
End Function

Private Function MthNmDrszVbe(A As Vbe, Optional WhStr$) As DRs
Set MthNmDrszVbe = DRs(MthNmFny, MthNmDryzVbe(A, WhStr))
End Function

Function MthNmDrszPj(A As VBProject, Optional WhStr$)
Set MthNmDrszPj = DRs(MthNmFny, MthNmDryzPj(A, WhStr))
End Function

Private Function MthNmDryzMd(M As CodeModule, Optional B As WhMth) As Variant()
MthNmDryzMd = DryAddCol3C(MthNmDryzSrc(Src(M), B), MdNm(M), ShtCmpTy(M.Parent.Type), PjNmzMd(M))
End Function

Private Function MthNmDryzVbe(A As Vbe, Optional WhStr$) As Variant()
Dim P
For Each P In PjItr(A, WhStr)
    PushIAy MthNmDryzVbe, MthNmDryzPj(CvPj(P), WhStr)
Next
End Function

Private Function MthNmDryzPj(P As VBProject, Optional WhStr$) As Variant()
Dim M, W As WhMth
Set W = WhMthzStr(WhStr)
For Each M In MdItrzPj(P, WhStr)
    PushIAy MthNmDryzPj, MthNmDryzMd(CvMd(M), W)
Next
End Function

Private Function MthNmDryzSrc(Src$(), Optional B As WhMth) As Variant()
Dim L
For Each L In Itr(Src)
    PushISomSz MthNmDryzSrc, MthNm3(L, B).MthNmDr
Next
End Function

