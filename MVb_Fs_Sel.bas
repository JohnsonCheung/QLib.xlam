Attribute VB_Name = "MVb_Fs_Sel"
Option Explicit

Function FfnSel$(Ffn$, Optional FSpec$ = "*.*", Optional Tit$ = "Select a file", Optional BtnNm$ = "Use the File Name")
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

Function PthSel$(Pth, Optional Tit$ = "Select a Path", Optional BtnNm$ = "Use this path")
With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
    .InitialFileName = IIf(IsNull(Pth), "", Pth)
    .Show
    If .SelectedItems.Count = 1 Then
        PthSel = EnsPthSfx(.SelectedItems(1))
    End If
End With
End Function

Private Sub Z_PthSel()
GoTo ZZ
ZZ:
MsgBox FfnSel("C:\")
End Sub


Private Sub Z()
Z_PthSel
MVb_Fs_Sel:
End Sub
