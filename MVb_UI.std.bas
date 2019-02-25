Attribute VB_Name = "MVb_UI"
Option Explicit
Function Cfm(Msg$) As Boolean
Cfm = MsgBox(Msg, vbYesNo + vbDefaultButton2) = vbYes
End Function
Function CfmYes(Msg$) As Boolean
CfmYes = UCase(InputBox(Msg)) = "YES"
End Function

Sub PromptCnl(Msg)
If MsgBox(Msg, vbOKCancel) = vbCancel Then Stop
End Sub
