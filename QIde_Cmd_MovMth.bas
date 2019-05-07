Attribute VB_Name = "QIde_Cmd_MovMth"
Option Explicit
Private Const CMod$ = "MIde_Cmd_MovMth."
Private Const Asm$ = "QIde"
Const BarNmzMovMth$ = "XMov_Mth"
Const BtnNmzMovMth$ = "XMov_Mth"

Property Get BarNy() As String()
BarNy = Itn(Bars)
End Property

Private Sub Z_Mov_MthBar()
MsgBox BarzMovMth.Name
End Sub

Function BarszVbe(A As Vbe) As Office.CommandBars
Set BarszVbe = A.CommandBars
End Function

Property Get Bars() As Office.CommandBars
Set Bars = BarszVbe(CurVbe)
End Property

Function HasBar(BarNm$) As Boolean
HasBar = HasItn(Bars, BarNm)
End Function
Function Bar(BarNm$) As Office.CommandBar
Set Bar = Bars(BarNm)
End Function
Sub RmvBar(BarNm$)
If HasBar(BarNm) Then Bars(BarNm).Delete
End Sub
Function HasBtn(A As Office.CommandBar, BtnCaption$) As Boolean
Dim C As Office.CommandBarControl
For Each C In A.Controls
    If C.Type = msoControlButton Then
        If CvBtn(C).Caption = BtnCaption Then HasBtn = True: Exit Function
    End If
Next
End Function
Sub EnsBarBtn(BarNm$, BtnCaption$)
EnsBar BarNmzMovMth
If HasBtn(Bars(BarNm), BtnCaption) Then Exit Sub
Bars(BarNm).Controls.Add(msoControlButton).Caption = BtnCaption
End Sub
Sub EnsBar(BarNm$)
If HasBar(BarNm) Then Exit Sub
AddBar BarNm
End Sub
Sub AddBar(BarNm$)
Bars.Add BarNm
End Sub
Private Property Get BarzMovMth() As Office.CommandBar
Set BarzMovMth = Bars(BarNmzMovMth)
End Property
Private Property Get BtnzMovMth() As Office.CommandBarControl
Set BtnzMovMth = BarzMovMth.Controls(BtnNmzMovMth)
End Property

Private Sub Z()
Z_Mov_MthBar
MIde_CMdMov_Mth:
End Sub
