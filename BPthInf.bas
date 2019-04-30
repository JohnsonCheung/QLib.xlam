Attribute VB_Name = "BPthInf"
Function HasSiblingFdr(Pth$, Fdr$) As Boolean
HasSiblingFdr = HasFdr(ParPth(Pth), Fdr)
End Function

Function SiblingPth$(Pth$, SiblingFdr$)
SiblingPth = AddFdrEns(ParPth(Pth), SiblingFdr)
End Function

