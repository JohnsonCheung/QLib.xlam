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
Sub ClsWinzWin(W As VBIDE.Window)
W.Close
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
Dim W, Exl(), Vis() As VBIDE.Window
Exl = ExlWinAp
For Each W In CurVbe.Windows
    If Not HasObj(Exl, W) Then
        ClsWinzWin CvWin(W)
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
If HasCmp(MdNm) Then Set WinzMdNm = Md(MdNm).CodePane.Window
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
EdtclrBtn.Execute
End Sub

Property Get WinCnt&()
WinCnt = CurVbe.Windows.Count
End Property

Function MdNmCdWin$(CdWin As VBIDE.Window)
MdNmCdWin = StrBet(CdWin.Caption, " - ", " (Code)")
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

Function SetVisWin(A As VBIDE.Window) As VBIDE.Window
A.Visible = True
A.WindowState = vbext_ws_Maximize
Set SetVisWin = A
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
Function CdPnezCmpNm(CmpNm) As CodePane
Set CdPnezCmpNm = Md(CmpNm).CodePane
End Function
Function WinzCmpNm(CmpNm) As VBIDE.Window
Set WinzCmpNm = CdPnezCmpNm(CmpNm).Window
End Function
Sub ShwCmp(CmpNm)
Dim C As VBIDE.CodePane: Set C = CdPnezCmpNm(CmpNm)
If IsNothing(C) Then Debug.Print "No such WinzCmpNm": Exit Sub
C.Show
End Sub

Sub ClsWinzExlCmpOoImm(MdNm$)
ClsWinzExl WinzImm, WinzMdNm(MdNm)
End Sub

Function WinAyAv(Av()) As VBIDE.Window()
Dim I
For Each I In Itr(Av)
    PushObj WinAyAv, I
Next
End Function

Sub ClsWinzExlWinAp(ParamArray WinAp())
Dim W, Av(): Av = WinAp
For Each W In Itr(VisWinAy)
    If Not HasObj(Av, W) Then
        CvWin(W).Close
    Else
        ShwWin CvWin(W)
    End If
Next
End Sub
Sub ShwWin(A As VBIDE.Window)
A.Visible = True
End Sub


Property Get VisWinAy() As VBIDE.Window()
Dim W As VBIDE.Window
For Each W In CurVbe.Windows
    If W.Visible Then PushObj VisWinAy, W
Next
End Property
