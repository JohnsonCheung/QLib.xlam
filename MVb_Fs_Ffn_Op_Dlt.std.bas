Attribute VB_Name = "MVb_Fs_Ffn_Op_Dlt"
Option Explicit

Sub DltFfnyAyIf(FfnAy)
Dim F
For Each F In Itr(FfnAy)
    DltFfnIf F
Next
End Sub

Sub DltFfn(A)
On Error GoTo X
Kill A
Exit Sub
X:
Thw CSub, "Cannot kill", "Ffn Er", A, Err.Description
End Sub

Sub DltFfnIf(A)
If HasFfn(A) Then DltFfn A
End Sub

