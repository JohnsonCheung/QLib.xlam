VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MthCnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public MdNm$, NPubSub%, NPubFun%, NPubPrp%, NPrvSub%, NPrvFun%, NPrvPrp%, NFrdSub%, NFrdFun%, NFrdPrp%
Friend Function Init(MdNm$, NPubSub%, NPubFun%, NPubPrp%, NPrvSub%, NPrvFun%, NPrvPrp%, NFrdSub%, NFrdFun%, NFrdPrp%) As MthCnt
With Me
.MdNm = MdNm
.NPubSub = NPubSub
.NPubFun = NPubFun
.NPubPrp = NPubPrp
.NPrvSub = NPrvSub
.NPrvFun = NPrvFun
.NPrvPrp = NPrvPrp
.NFrdSub = NFrdSub
.NFrdFun = NFrdFun
.NFrdPrp = NFrdPrp
End With
Set Init = Me
End Function
Property Get N%()
N = NPubSub + NPubFun + NPubPrp + NPrvSub + NPrvFun + NPrvPrp + NFrdSub + NFrdFun + NFrdPrp
End Property
Function Lin$(Optional Hdr As eHdr)
Dim Pfx$: If Hdr = eeWithHdr Then Pfx = "Pub* | Prv* | Frd* : *{Sub Fun Frd} "
Lin = Pfx & MdNm & " | " & N & " | " & NPubSub & " " & NPubFun & " " & NPubPrp & " | " & NPrvSub & " " & NPrvFun & " " & NPrvPrp & " | " & NFrdSub & " " & NFrdFun & " " & NFrdPrp
End Function

