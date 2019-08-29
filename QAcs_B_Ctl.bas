Attribute VB_Name = "QAcs_B_Ctl"
Option Explicit
Option Compare Text

Sub SetNoTabStop(A As Access.Form)
DoItrFun A.Controls, "CmdTurnOffTabStop"
End Sub

Function CvAcsCtl(A) As Access.Control
Set CvAcsCtl = A
End Function

Function CvAcsBtn(A) As Access.CommandButton
Set CvAcsBtn = A
End Function

Function CvAcsTgl(A) As Access.ToggleButton
Set CvAcsTgl = A
End Function

Sub SetTBox(A As Access.TextBox, Msg$)
Dim CrLf$, B$
If A.Value <> "" Then CrLf = vbCrLf
B = LineszLasN(A.Value & CrLf & Now & " " & Msg, 5)
A.Value = B
DoEvents
End Sub

Sub PthzSelzTxtb(A As Access.TextBox)
Dim R$
R = PthzSel(A.Value)
If R = "" Then Exit Sub
A.Value = R
End Sub
Sub TurnOffTabStop(A As Access.Control)
If Not HasPfx(A.Name, "Cmd") Then Exit Sub
Select Case True
Case IsBtn(A): CvBtn(A).TabStop = False
Case IsTglBtn(A): CvAcsTgl(A).TabStop = False
End Select
End Sub


'
