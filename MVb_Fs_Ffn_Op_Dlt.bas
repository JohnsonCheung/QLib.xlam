Attribute VB_Name = "MVb_Fs_Ffn_Op_Dlt"
Option Explicit

Sub DltFfnyAyIf(FfnSy$())
Dim Ffn
For Each Ffn In Itr(FfnSy)
    DltFfnIf CStr(Ffn)
Next
End Sub

Sub DltFfn(Ffn$)
On Error GoTo X
Kill Ffn
Exit Sub
X:
Thw CSub, "Cannot kill", "Ffn Er", Ffn, Err.Description
End Sub

Sub DltFfnIf(Ffn$)
If HasFfn(Ffn$) Then DltFfn Ffn
End Sub

Function DltFfnIfPrompt(Ffn$, Msg$) As Boolean 'Return true if error
If Not HasFfn(Ffn$) Then Exit Function
On Error GoTo X
Kill Ffn
Exit Function
X:
MsgBox "File [" & Ffn & "] cannot be deleted, " & vbCrLf & Msg
DltFfnIfPrompt = True
End Function

Function DltFfnDone(Ffn$) As Boolean
On Error GoTo X
Kill Ffn
DltFfnDone = True
Exit Function
X:
End Function
