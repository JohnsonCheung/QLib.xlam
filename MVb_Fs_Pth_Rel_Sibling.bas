Attribute VB_Name = "MVb_Fs_Pth_Rel_Sibling"
Option Explicit
Function HasSiblingFdr(Pth, Fdr) As Boolean
HasSiblingFdr = HasFdr(ParPth(Pth), Fdr)
End Function

Function SiblingPth$(Pth, SiblingFdr)
SiblingPth = AddFdrEns(ParPth(Pth), SiblingFdr)
End Function
