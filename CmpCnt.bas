VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CmpCnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public NMod%, NCls%, NDoc%, NOth%, Locked As Boolean
Enum eHdr
    eeNoHdr
    eeWithHdr
End Enum
Friend Function Init(NMod%, NCls%, NDoc%, NOth%) As CmpCnt
With Me
    .NMod = NMod
    .NCls = NCls
    .NDoc = NDoc
    .NOth = NOth
End With
Set Init = Me
End Function

Property Get NCmp%()
NCmp = NMod + NCls + NDoc + NOth
End Property

Function Lin$(Optional Hdr As eHdr)
Dim Pfx$
If Hdr = eeWithHdr Then Pfx = "Cmp Mod Cls Doc Oth "
Lin = Pfx & NCmp & " " & NMod & " " & NCls & " " & NDoc & " " & NOth
End Function
