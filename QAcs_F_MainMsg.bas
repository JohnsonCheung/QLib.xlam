Attribute VB_Name = "QAcs_F_MainMsg"
Option Explicit
Option Compare Text
'Assume there is Application.Forms("Main").Msg (TextBox)
'MMsg means Main.Msg (TextBox)
Sub ClrMainMsg()
SetMainMsg ""
End Sub

Sub SetMainMsgzQnm(QryNm)
SetMainMsg "Running query: (" & QryNm & ")...."
End Sub

Private Sub SetMainMsg(Msg$)
On Error Resume Next
SetTBox MainMsgBox, Msg
End Sub

Private Property Get MainMsgBox() As Access.TextBox
On Error Resume Next
Set MainMsgBox = MainFrm.Controls("Msg")
End Property

Private Property Get MainFrm() As Access.Form
On Error Resume Next
Set MainFrm = Access.Forms("Main")
End Property


'
