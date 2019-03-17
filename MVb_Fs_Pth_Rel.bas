Attribute VB_Name = "MVb_Fs_Pth_Rel"
Option Explicit

Function ParPth$(Pth)
If Not HasSubStr(Pth, PthSep) Then Err.Raise 1, "ParPth", "No PthSep in Pth" & vbCrLf & Pth
ParPth = StrBefRevOrAll(RmvLasChr(PthEnsSfx(Pth)), PthSep) & PthSep
End Function

Function ParFdr$(Pth)
ParFdr = Fdr(ParPth(Pth))
End Function

Function ParPthN$(Pth, UpN%)
Dim O$, J%
O = Pth
For J = 1 To UpN
    O = ParPth(O)
Next
ParPthN = O
End Function

