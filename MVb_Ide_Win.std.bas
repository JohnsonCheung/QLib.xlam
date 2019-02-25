Attribute VB_Name = "MVb_Ide_Win"
Option Explicit
Const CMod$ = "MVb_Ide_Z_Win."

Property Get CdWinAy() As VBIDE.Window()
CdWinAy = WinAyWinTy(vbext_wt_CodeWindow)
End Property

Sub ClrWinzImm()
DoEvents
With WinzImm
    .SetFocus
    .Visible = True
End With
SndKeys "^{HOME}^+{END}"
DoEvents
'SndKeys "{DEL}" '<-- it does not work?
'DoEvents
End Sub

Sub ClsWin()
Dim W As VBIDE.Window
For Each W In CurVbe.Windows
    If W.Visible Then W.Close
Next
End Sub

Sub ClsWinzExlWin(A As VBIDE.Window)
ClsWinzExl A
End Sub

Sub ClsWinzExl(ParamArray ExlWinAp())
Stop
Dim W, Exl(), Vis() As VBIDE.Window
'V = VisWinAy
Exl = ExlWinAp
For Each W In Itr(Vis)
    If Not HasObj(Exl, W) Then
        ClsWinOpt W
    Else
        ShwWinOpt W
    End If
Next
End Sub
Sub SetWinVisOpt(A, Vis As Boolean)
Dim W As VBIDE.Window
Set W = CvWinOpt(A)
If Not IsNothing(W) Then W.Visible = Vis
End Sub
Sub ClsWinOpt(A)
SetWinVisOpt A, False
End Sub
Sub ShwWinOpt(A)
SetWinVisOpt A, True
End Sub
Function IsWin(A) As Boolean
IsWin = TypeName(A) = "Window"
End Function
Sub ClsWinzExlMd(ExlMdNm$)
ClsWinzExlWin WinzMd(Md(ExlMdNm))
End Sub

Sub ClsWinzImm()
DoEvents
WinzImm.Visible = False
End Sub

Sub ClsWinzExlImm()
ClsWinzExlWin WinzImm
End Sub

Property Get CurCdWin() As VBIDE.Window
On Error Resume Next
Set CurCdWin = CurVbe.ActiveCodePane.Window
End Property

Private Property Get CurVbe() As Vbe
Set CurVbe = Application.Vbe
End Property

Property Get CurWin() As VBIDE.Window
Set CurWin = CurVbe.ActiveWindow
End Property

Function CvWinAy(A) As VBIDE.Window()
CvWinAy = A
End Function

Property Get EmpWinAy() As VBIDE.Window()
End Property

Property Get WinzImm() As VBIDE.Window
Set WinzImm = FstWinTy(vbext_wt_Immediate)
End Property

Property Get WinzLcl() As VBIDE.Window
Set WinzLcl = FstWinTy(vbext_wt_Locals)
End Property

Function WinzMdNm(MdNm) As VBIDE.Window
If HasCmp(MdNm) Then Set WinzMdNm = WinzMd(Md(MdNm))
End Function

Function WinzMd(A As CodeModule) As VBIDE.Window
Set WinzMd = A.CodePane.Window
End Function

Private Function MdzPj(A As VBProject, Nm) As CodeModule
Set MdzPj = A.VBComponents(Nm).CodeModule
End Function

Sub ShwDbg()
ClsWinzExl WinzImm, WinzLcl, CurCdWin
DoEvents
TileV
ClrWinzImm
End Sub

Sub JmpNxtStmt()
Const CSub$ = CMod & "JmpNxtStmt"
With JmpNxtStmtBtn
    If Not .Enabled Then
        'Msg CSub, "JmpNxtStmtBtn is disabled"
        Exit Sub
    End If
    .Execute
End With
End Sub

Property Get VisWinCnt&()
VisWinCnt = NItrPrpTrue(CurVbe.Windows, "Visible")
End Property
Function CvWin(A) As VBIDE.Window
Set CvWin = A
End Function
Function CvWinOpt(A) As VBIDE.Window
On Error Resume Next
Set CvWinOpt = A
End Function
Sub ClrWin(A As VBIDE.Window)
DoEvents
SelAllBtn.Execute
DoEvents
SendKeys " "
EdtClrBtn.Execute
End Sub

Property Get WinCnt&()
WinCnt = CurVbe.Windows.Count
End Property

Function MdNmCdWin$(CdWin As VBIDE.Window)
MdNmCdWin = TakBet(CdWin.Caption, " - ", " (Code)")
End Function

Property Get WinNy() As String()
Dim W As VBIDE.Window
For Each W In CurVbe.Windows
    Debug.Print W.Caption, W.Type
    PushI WinNy, W.Caption
Next
End Property

Function FstWinTy(A As vbext_WindowType) As VBIDE.Window
Set FstWinTy = FstItrPEv(CurVbe.Windows, "Type", A)
End Function

Function WinAyWinTy(T As vbext_WindowType) As VBIDE.Window()
WinAyWinTy = ItrwPrpEqval(CurVbe.Windows, "Type", T)
End Function

Function SetWinzis(A As VBIDE.Window) As VBIDE.Window
A.Visible = True
A.WindowState = vbext_ws_Maximize
Set SetWinzis = A
End Function

Private Sub Z_Md()
Debug.Print Md("MVb_Ide_Z_Win").Parent.Name
End Sub

Private Sub ZZ()
Dim A()
Dim B$
Dim C As vbext_WindowType
Dim D As Variant
Dim E As CodeModule
Dim F As VBIDE.Window

ClrWinzImm

CvWin D
CvWinAy D
WinzMd E
ShwDbg
JmpNxtStmt
ClrWin F
MdNmCdWin F
End Sub

Private Sub Z()
Z_Md
End Sub
Function CmpWin(CmpNm) As VBIDE.Window
Dim W As VBIDE.Window
For Each W In CurVbe.Windows
    If W.Type = vbext_wt_CodeWindow Then
        If IsWinCaptionEqMd(W.Caption, CmpNm) Then Set CmpWin = W: Exit Function
    End If
Next
End Function
Function IsWinCaptionEqMd(WinCaption$, CmpNm) As Boolean
IsWinCaptionEqMd = TakBet(WinCaption, " - ", " (Code)") = CmpNm
End Function
Sub ShwCmp(CmpNm)
Dim W As VBIDE.Window: Set W = CmpWin(CmpNm)
If IsNothing(W) Then Debug.Print "No such CmpWin": Exit Sub
W.Visible = True
End Sub

Sub ClsWinzExptMdImm(MdNm$)
ClsWinzExpt WinzImm, WinzMdNm(MdNm)
End Sub

Function WinAyAv(Av()) As VBIDE.Window()
Dim I
For Each I In Itr(Av)
    PushObj WinAyAv, I
Next
End Function

Sub ClsWinzExpt(ParamArray WinAp())
Dim W, Av(): Av = WinAp
For Each W In Itr(VisWinAy)
    If Not HasObj(Av, W) Then
        CvWin(W).Close
    Else
        ShwWinOpt W
    End If
Next
End Sub

Sub ShwWin(ParamArray WinAp())
Dim I, W As VBIDE.Window
For Each I In Av
    Set W = CvWinOpt(I)
    If Not IsNothing(W) Then
        W.Visible = True
    End If
Next
End Sub


Property Get VisWinAy() As VBIDE.Window()
Dim W As VBIDE.Window
For Each W In CurVbe.Windows
    If W.Visible Then PushObj VisWinAy, W
Next
End Property
