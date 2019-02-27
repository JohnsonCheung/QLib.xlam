Attribute VB_Name = "MIde_Cmd_Wh"
Option Explicit

Private Function WhMthzStrPfx$(Pfx, IncPrv As Boolean)

End Function
Private Function WhMthzStrSfx$(Sfx, IncPrv As Boolean)

End Function
Private Function WhMthzStrPatn$(Patn, IncPrv As Boolean)

End Function

Sub LisMth(Pfx$, Optional InclPrv As Boolean)
LisMthPfx Pfx, InclPrv
End Sub

Sub LisMthPfx(Pfx$, Optional InclPrv As Boolean)
D MthDNy(WhMthzStrPfx(Pfx, InclPrv))
End Sub

Sub LisMthSfx(Sfx$, Optional InclPrv As Boolean)
D MthDNy(WhMthzStrSfx(Sfx, InclPrv))
End Sub

Sub LisMthPatn(Patn$, Optional InclPrv As Boolean)
D MthNyPj(WhMthzStrPatn(Patn, InclPrv))
End Sub

