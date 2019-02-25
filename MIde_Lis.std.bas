Attribute VB_Name = "MIde_Lis"
Option Explicit
Function MdzPjLisDt(A As VBProject, Optional B As WhMd) As Dt
Stop '
End Function

Sub MdzPjLisBrwDt(A As VBProject, Optional B As WhMd)
BrwDt MdzPjLisDt(A, B)
End Sub

Sub MdzPjLisDmpDt(A As VBProject, Optional B As WhMd)
DmpDt MdzPjLisDt(A, B)
End Sub

Sub LisMd(Optional Patn$, Optional Exl$)
Dim A$()
'    A = CmpNyPj(CurPj, WhMd("Std", WhNm(Patn, Exl)))
    A = AySrt(A)
    A = AyAddPfx(A, "ShwMbr """)
D A
End Sub
Sub LisPj()
Dim A$()
    A = PjNyz(CurVbe)
    D AyAddPfx(A, "ShwPj """)
D A
End Sub

Sub LisMth(Optional MthPatn$, Optional MthExl$, Optional WhMdy$, Optional WhKd$, Optional MdPatn$)
Dim Ny$(), M As WhMdMth
'    Set M = NewWhMdMth_MTH_MD(MthPatn, MthExl, WhMdy, WhKd, MdPatn)
'    Ny = MthNyzPj(CurPj, M)
D AyAddPfx(Ny, PjNm & ".")
End Sub


