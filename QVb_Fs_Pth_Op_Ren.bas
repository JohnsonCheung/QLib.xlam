Attribute VB_Name = "QVb_Fs_Pth_Op_Ren"
Option Explicit
Private Const CMod$ = "MVb_Fs_Pth_Op_Ren."
Private Const Asm$ = "QVb"
Sub RenPthAddPfx(Pth, Pfx)
RenPth Pth, AddPfxzPth(Pth, Pfx)
End Sub

Sub RenPth(Pth, NewPth)
If HasPth(NewPth) Then Thw CSub, "NewPth Has", "Pth NewPth", Pth, NewPth
If Not HasPth(Pth) Then Thw CSub, "Pth not Has", "Pth NewPth", Pth, NewPth
Fso.GetFolder(Pth).Name = NewPth
End Sub

