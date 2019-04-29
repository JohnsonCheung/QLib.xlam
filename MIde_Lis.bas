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
    A = SyAddPfx(A, "ShwMbr """)
D A
End Sub
Sub LisPj()
Dim A$()
    A = PjNyzVbe(CurVbe)
    D SyAddPfx(A, "ShwPj """)
D A
End Sub

Sub LisStopLin()

End Sub
Sub LisMth(Optional WhStr$)
Dim Ay$(): Ay = MthQNyzVbe(CurVbe, WhStr)
Debug.Print "Fst 30 of " & Si(Ay) & " methods"
D AywFstNEle(Ay, 30)
End Sub

Private Function WhStrzMthPatn$(MthPatn$, Optional PubOnly As Boolean)
WhStrzMthPatn = " -MthPatn " & MthPatn & WhStrzPubOnly(PubOnly)
End Function

Private Function WhStrzPubOnly$(PubOnly As Boolean)
If PubOnly Then WhStrzPubOnly = " -Pub"
End Function

Function WhStrzPfx$(MthPfx$, Optional PubOnly As Boolean)
WhStrzPfx = WhStrzMthPatn("^" & MthPfx, PubOnly)
End Function

Function WhStrzSfx$(MthSfx$, Optional PubOnly As Boolean)
WhStrzSfx = WhStrzMthPatn(MthSfx & "$", PubOnly)
End Function

Sub LisMthPfx(Pfx$, Optional PubOnly As Boolean)
D MthQNyInVbe(WhStrzPfx(Pfx, PubOnly))
End Sub

Sub LisMthSfx(Sfx$, Optional PubOnly As Boolean)
D MthQNyInVbe(WhMthzSfx(Sfx, PubOnly))
End Sub

Sub LisMthPatn(Patn$, Optional InclPrv As Boolean)
D MthQLyInVbe(WhStrzMthPatn(Patn, InclPrv))
End Sub
