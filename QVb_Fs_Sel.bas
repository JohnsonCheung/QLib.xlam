Attribute VB_Name = "QVb_Fs_Sel"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Fs_Sel."
Private Const Asm$ = "QVb"

Function FfnSel$(Ffn, Optional FSpec$ = "*.*", Optional Tit$ = "Select a file", Optional BtnNm$ = "Use the File Name")
With Application.FileDialog(msoFileDialogFilePicker)
    .Filters.Clear
    .Title = Tit
    .AllowMultiSelect = False
    .Filters.Add "", FSpec
    .InitialFileName = Ffn
    .ButtonName = BtnNm
    .Show
    If .SelectedItems.Count = 1 Then
        FfnSel = .SelectedItems(1)
    End If
End With
End Function

Function SelPth$(Pth, Optional Tit$ = "Select a Path", Optional BtnNm$ = "Use this path")
With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
    .InitialFileName = IIf(IsNull(Pth), "", Pth)
    .Show
    If .SelectedItems.Count = 1 Then
        SelPth = EnsPthSfx(.SelectedItems(1))
    End If
End With
End Function

Private Sub Z_SelPth()
GoTo ZZ
ZZ:
MsgBox FfnSel("C:\")
End Sub


Private Sub ZZ()
Z_SelPth
MVb_Fs_Sel:
End Sub
