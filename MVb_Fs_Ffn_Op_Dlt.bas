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

Sub DltFfnIf(Ffn)
If HasFfn(Ffn) Then DltFfn Ffn
End Sub

Function DltFfnIfPrompt(Ffn, Msg$) As Boolean 'Return true if error
If Not HasFfn(Ffn) Then Exit Function
On Error GoTo X
Kill Ffn
Exit Function
X:
MsgBox "File [" & Ffn & "] cannot be deleted, " & vbCrLf & Msg
DltFfnIfPrompt = True
End Function
