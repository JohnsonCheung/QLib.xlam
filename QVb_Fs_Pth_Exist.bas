Attribute VB_Name = "QVb_Fs_Pth_Exist"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Fs_Pth_Exist."
Private Const Asm$ = "QVb"
Function EnsPth$(Pth)
Dim P$: P = EnsPthSfx(Pth)
If Not Fso.FolderExists(Pth) Then MkDir RmvLasChr(P)
EnsPth = Pth
End Function

Sub EnsPthzAllSeg(Pth)
Dim J%, O$, Ay$()
Ay = Split(RmvSfx(Pth, PthSep), PthSep)
O = Ay(0)
For J = 1 To UBound(Ay)
    O = O & PthSep & Ay(J)
    EnsPth O
Next
End Sub

Function HasPth(Pth) As Boolean
HasPth = IsPthExist(Pth)
End Function

Function HasFdr(Pth, Fdr$) As Boolean
HasFdr = HasEle(FdrAy(Pth), Fdr)
End Function

Sub ThwIf_PthNotExist(Pth, Fun$)
If Not HasPth(Pth) Then Thw Fun, "Pth not exist", "Pth", Pth
End Sub

Function AnyFil(Pth) As Boolean
AnyFil = Dir(Pth) <> ""
End Function
Function IsPth(Pth) As Boolean
IsPth = IsPthExist(Pth)
End Function
Function IsPthExist(Pth) As Boolean
IsPthExist = Fso.FolderExists(Pth)
End Function

Function HasSubFdr(Pth) As Boolean
HasSubFdr = Fso.GetFolder(Pth).SubFolders.Count > 0
End Function

Sub ThwIf_PthNotExist1(Pth, Optional Fun$ = "ThwIf_PthNotExist1")
If Not HasPth(Pth) Then Thw Fun, "Path not exist", "Path", Pth
End Sub


