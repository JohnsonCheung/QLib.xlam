Attribute VB_Name = "MVb_Fs_Pth_Op_Ren"
Option Explicit
Sub RenPthAddPfx(Pth, Pfx)
RenPth Pth, PthAddPfx(Pth, Pfx)
End Sub

Sub RenPth(Pth, NewPth)
If HasPth(NewPth) Then Thw CSub, "NewPth Has", "Pth NewPth", Pth, NewPth
If Not HasPth(Pth) Then Thw CSub, "Pth not Has", "Pth NewPth", Pth, NewPth
Fso.GetFolder(Pth).Name = NewPth
End Sub

