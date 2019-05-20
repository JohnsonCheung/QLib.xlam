Attribute VB_Name = "QVb_Fs_Ffn_Op_Mov"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Fs_Ffn_Op_Mov."
Private Const Asm$ = "QVb"


Sub MovFilUp(Pth)
Dim I, Tar$
Tar$ = ParPth(Pth)
For Each I In Itr(FnAy(Pth))
    MovFfn CStr(I), Tar
Next
End Sub


Sub MovFfn(Ffn, ToPth$)
Fso.MoveFile Ffn, ToPth
End Sub


