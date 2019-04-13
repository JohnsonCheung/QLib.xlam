Attribute VB_Name = "MIde_Cmd_MovMth"
Option Explicit
Const XMov_MthBarNm$ = "XMov_Mth"
Const XMov_MthBtnNm$ = "XMov_Mth"

Property Get CmdBarNy() As String()
CmdBarNy = Itn(CurVbe_Bars)
End Property

Private Sub Z_Mov_MthBar()
MsgBox XMov_MthBar.Name
End Sub

Function Vbe_Bars(A As Vbe) As Office.CommandBars
Set Vbe_Bars = A.CommandBars
End Function

Property Get CurVbe_Bars() As Office.CommandBars
Set CurVbe_Bars = Vbe_Bars(CurVbe)
End Property

Function CurVbe_BarsHas(A) As Boolean
CurVbe_BarsHas = HasItn(CurVbe_Bars, A)
End Function
Function CmdBar(A) As Office.CommandBar
Set CmdBar = CurVbe_Bars(A)
End Function
Sub RmvCmdBar(A)
If CurVbe_BarsHas(A) Then CmdBar(A).Delete
End Sub
Function CmdBar_HasBtn(A As Office.CommandBar, BtnCaption)
Dim C As Office.CommandBarControl
For Each C In A.Controls
    If C.Type = msoControlButton Then
        If CvCmdBtn(C).Caption = BtnCaption Then CmdBar_HasBtn = True: Exit Function
    End If
Next
End Function
Sub EnsCmdBarBtn(CmdBarNm, BtnCaption)
EnsCmdBar XMov_MthBarNm
If CmdBar_HasBtn(CmdBar(CmdBarNm), BtnCaption) Then Exit Sub
CmdBar(CmdBarNm).Controls.Add(msoControlButton).Caption = BtnCaption
End Sub
Sub EnsCmdBar(A$)
If CurVbe_BarsHas(A) Then Exit Sub
AddCmdBar A
End Sub
Sub AddCmdBar(A)
CurVbe_Bars.Add A
End Sub
Private Property Get XMov_MthBar() As Office.CommandBar
Set XMov_MthBar = CurVbe_Bars(XMov_MthBarNm)
End Property
Private Property Get XMov_MthBtn() As Office.CommandBarControl
Set XMov_MthBtn = XMov_MthBar.Controls(XMov_MthBtnNm)
End Property

Private Sub Z()
Z_Mov_MthBar
MIde_CMdMov_Mth:
End Sub
