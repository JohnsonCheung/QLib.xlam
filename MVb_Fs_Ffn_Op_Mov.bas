Attribute VB_Name = "MVb_Fs_Ffn_Op_Mov"
Option Explicit


Sub MovFilUp(Pth)
Dim I, Tar$
Tar$ = ParPth(Pth)
For Each I In Itr(FnSy(Pth))
    MovFfn CStr(I), Tar
Next
End Sub


Sub MovFfn(Ffn$, ToPth$)
Fso.MoveFile Ffn, ToPth
End Sub


