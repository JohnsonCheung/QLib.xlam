Attribute VB_Name = "QVb_Ide_Win"
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Ide_Win."

Sub ClrImm()
DoEvents
With WinOfImm
    .SetFocus
    .Visible = True
End With
SndKeys "^{HOME}^+{END}"
DoEvents
'SndKeys "{DEL}" '<-- it does not work?
'DoEvents
End Sub
ClsWin

Sub ClsWin()
Dim W As VBIDE.Window
For Each W In CVbe.Windows
    If W.Visible Then W.Close
Next
End Sub

Function IsWin(A) As Boolean
IsWin = TypeName(A) = "Window"
End Function

Function CvWiny(A) As VBIDE.Window()
CvWiny = A
End Function

Property Get EmpWiny() As VBIDE.Window()
End Property

Property Get WinOfImm() As VBIDE.Window
Set WinOfImm = FstWinTy(vbext_wt_Immediate)
End Property

Property Get WinOfLcl() As VBIDE.Window
Set WinOfLcl = FstWinTy(vbext_wt_Locals)
End Property

Function WinOfMdn(Mdn) As VBIDE.Window
If HasCmpzN(Mdn) Then Set WinOfMdn = Md(Mdn).CodePane.Window
End Function

Function WinzM(A As CodeModule) As VBIDE.Window
Set WinzM = A.CodePane.Window
End Function

Private Function MdzP(P As VBProject, Nm) As CodeModule
Set MdzP = P.VBComponents(Nm).CodeModule
End Function

Sub ShwDbg()
ClsWinOfExl WinOfImm, WinOfLcl, CurCdWin
DoEvents
TileV
ClrWinOfImm
End Sub

Sub JmpNxtStmt()
Const CSub$ = CMod & "JmpNxtStmt"
With BtnOfJmpNxtStmt
    If Not .Enabled Then
        'Msg CSub, "BtnOfJmpNxtStmt is disabled"
        Exit Sub
    End If
    .Execute
End With
End Sub

Property Get VisWinCnt&()
VisWinCnt = NItrPrpTrue(CVbe.Windows, "Visible")
End Property
Function CvWin(A) As VBIDE.Window
Set CvWin = A
End Function

Sub ClrWin(A As VBIDE.Window)
DoEvents
BtnOfSelAll.Execute
DoEvents
SendKeys " "
BtnOfEdtClr.Execute
End Sub

Property Get WinCnt&()
WinCnt = CVbe.Windows.Count
End Property

Function MdnCdWin$(CdWin As VBIDE.Window)
MdnCdWin = Bet(CdWin.Caption, " - ", " (Code)")
End Function

Property Get WinNy() As String()
Dim W As VBIDE.Window
For Each W In CVbe.Windows
    Debug.Print W.Caption, W.Type
    PushI WinNy, W.Caption
Next
End Property

Function FstWinTy(A As vbext_WindowType) As VBIDE.Window
Set FstWinTy = FstItmPEv(CVbe.Windows, PrpPth("Type"), A)
End Function

Function WinyWinTy(T As vbext_WindowType) As VBIDE.Window()
WinyWinTy = ItrwPrpEqval(CVbe.Windows, "Type", T)
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

ClrWinOfImm

CvWin D
CvWiny D
WinOfMd E
ShwDbg
JmpNxtStmt
ClrWin F
MdnCdWin F
End Sub

Function PnezCmpn(Cmpn$) As CodePane
Set PnezCmpNm = Md(Cmpn).CodePane
End Function
Function WinOfCmpn(Cmpn$) As VBIDE.Window
Set WinOfCmpn = PnezCmpn(Cmpn).Window
End Function
Sub JmpCmp(Cmpn$)
Dim C As VBIDE.CodePane: Set C = CdPnezCmpNm(Cmpn)
If IsNothing(C) Then Debug.Print "No such WinOfCmpNm": Exit Sub
C.Show
End Sub

Sub ClsWinOfExl(ExlMdn$)
ClsWinOfExlAp WinOfImm, WinOfMdn(Mdn)
End Sub

Function WinyAv(Av()) As VBIDE.Window()
Dim I
For Each I In Itr(Av)
    PushObj WinyAv, I
Next
End Function

Sub ClsWinOfExlAp(ParamArray ExlWinAp())
Dim I, W As VBIDE.Window, Av(): Av = WinAp
For Each I In Itr(VisWiny)
    Set W = I
    If Not HasObj(Av, W) Then
        ClsWinOfW W
    Else
        ShwWin W
    End If
Next
End Sub
Sub ShwWin(A As VBIDE.Window)
A.Visible = True
End Sub
Sub ClsWinOfW(W As VBIDE.Window)
W.Visible = False
End Sub
Property Get VisWinCapAy() As String()
VisWinCapAy = SyzOyPrp(VisWiny, PrpPth("Caption"))
End Property
Property Get VisWiny() As VBIDE.Window()
Dim W As VBIDE.Window
For Each W In CVbe.Windows
    If W.Visible Then PushObj VisWiny, W
Next
End Property
Function LnozM&(M As CodeModule)
LnozM = RRCCzPne(PnezM(M)).R1
End Function
Function RRCCzPne(P As CodePane) As RRCC
If IsNothing(P) Then Exit Function
Dim R1&, R2&, C1&, C2&
P.GetSelection R1, R2, C1, C2
RRCCzPne = RRCC(R1, R2, C1, C2)
End Function
Function MthnzM$(M As CodeModule)
Dim K As vbext_ProcKind
MthnzM = M.ProcOfLine(LnozM(M), K)
End Function

Function PnezM(M As CodeModule) As VBIDE.CodePane
Set PnezM = M.CodePane.ActiveCodePane
End Function

