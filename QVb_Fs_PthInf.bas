Attribute VB_Name = "QVb_Fs_PthInf"
Option Explicit
Option Compare Text
Private Const CMod$ = "BPthInf."
Function HasSiblingFdr(Pth, Fdr$) As Boolean
HasSiblingFdr = HasFdr(ParPth(Pth), Fdr)
End Function

Function SiblingPth$(Pth, SiblingFdr$)
SiblingPth = AddFdrEns(ParPth(Pth), SiblingFdr)
End Function

