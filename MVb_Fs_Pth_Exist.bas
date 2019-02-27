Attribute VB_Name = "MVb_Fs_Pth_Exist"
Option Explicit
Function EnsPth$(Pth)
If Not Fso.FolderExists(Pth) Then MkDir Pth
EnsPth = PthEnsSfx(Pth)
End Function
Function PthEns$(Pth)
PthEns = EnsPth(Pth)
End Function

Function PthEnsAll$(A$)
Dim Ay$(): Ay = Split(A, PthSep)
Dim J%, O$
O = Ay(0)
For J = 1 To UB(Ay)
    O = O & PthSep & Ay(J)
    PthEns O
Next
PthEnsAll = A
End Function


Function IsPth(Pth) As Boolean
IsPth = Fso.FolderExists(Pth)
End Function

Function HasFdr(Pth, Fdr) As Boolean
HasFdr = HasEle(FdrAy(Pth), Fdr)
End Function


Sub ThwNotPth(Pth)
If Not IsPth(Pth) Then Err.Raise 1, , "Not Pth(" & Pth & ")"
End Sub


Function HasFilPth(Pth) As Boolean
HasFilPth = Fso.GetFolder(Pth).Files.Count > 0
End Function
Function HasPth(Pth) As Boolean
HasPth = Fso.FolderExists(Pth)
End Function

Function HasSubFdr(Pth) As Boolean
HasSubFdr = Fso.GetFolder(Pth).SubFolders.Count > 0
End Function

Sub ThwNotHasPth(Pth, Fun$)
If Not HasPth(Pth) Then Thw Fun, "Path not Has", "Path", Pth
End Sub


