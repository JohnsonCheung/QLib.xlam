Attribute VB_Name = "QIde_B_CtlInf"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Cmd_MovMth."
Private Const Asm$ = "QIde"

Property Get BarNy() As String()
BarNy = Itn(Bars)
End Property

Function BarszV(A As Vbe) As Office.CommandBars
Set BarszV = A.CommandBars
End Function

Property Get StdBar() As Office.CommandBar
Set StdBar = Bars("Standard")
End Property

Property Get BoEdtClr() As Office.CommandBarButton
Set BoEdtClr = FstCaption(EditPopup.Controls, "C&lear")
End Property

Property Get BarOfMnu() As CommandBar
Set BarOfMnu = BarOfMnuzV(CVbe)
End Property

Property Get BoSelAll() As Office.CommandBarButton
Set BoSelAll = FstCaption(EditPopup.Controls, "Select &All")
End Property

Function IsBtn(A) As Boolean
IsBtn = TypeName(A) = "CommandButton"
End Function

Property Get BoNxtStmt() As CommandBarButton
Set BoNxtStmt = PopupOfDbg.Controls("Show Next Statement")
End Property


Property Get PopupOfWin() As CommandBarPopup
Set PopupOfWin = BarOfMnu.Controls("Window")
End Property
Property Get BoTileH() As CommandBarButton
Set BoTileH = PopupOfWin.Controls("Tile &Horizontally")
End Property

Property Get BoTileV() As Office.CommandBarButton
Set BoTileV = PopupOfWin.Controls("Tile &Vertically")
End Property

Property Get BoXls() As Office.CommandBarControl
Set BoXls = StdBar.Controls(1)
End Property

Private Sub Z_DbgPop()
Dim A
Set A = PopupOfDbg
Stop
End Sub

Private Sub Z_BarOfMnu()
Dim A As CommandBar
Set A = BarOfMnu
Stop
End Sub


Private Property Get EditPopup() As Office.CommandBarPopup
Set EditPopup = FstCaption(BarOfMnu.Controls, "&Edit")
End Property

Property Get BoSav() As CommandBarButton
Set BoSav = BoSavzV(CVbe)
End Property

Property Get BoJmpNxtStmt() As CommandBarButton
Set BoJmpNxtStmt = PopupOfDbg.Controls("Show Next Statement")
End Property


Function BarOfMnuzV(A As Vbe) As CommandBar
Set BarOfMnuzV = A.CommandBars("Menu Bar")
End Function

Function BoSavzV(A As Vbe) As CommandBarButton
Dim I As CommandBarControl, S As Office.CommandBarControls
Set S = StdBarzV(A).Controls
For Each I In S
    If HasPfx(I.Caption, "&Sav") Then Set BoSav = I: Exit Function
Next
Stop
End Function

Function StdBarzV(A As Vbe) As Office.CommandBar
Dim X As Office.CommandBars
Set X = BarszV(A)
Set StdBarzV = X("Standard")
End Function

Function CvBtn(A) As Office.CommandBarButton
Set CvBtn = A
End Function


Function CvCtl(A) As CommandBarControl
Set CvCtl = A
End Function

Property Get PopupOfDbg() As CommandBarPopup
Set PopupOfDbg = BarOfMnu.Controls("Debug")
End Property

Private Function FstCaption(Itr, Caption$) 'Return FstItm with Caption-Prp = Caption$
FstCaption = FstItmPEv(Itr, PrpPth("Caption"), Caption)
End Function

Function BarNyzV(A As Vbe) As String()
BarNyzV = Itn(A.CommandBars)
End Function

Function CapNy(A As Controls) As String()
CapNy = SyzItrP(A, PrpPth("Caption"))
End Function

Property Get WinOfBrwObj() As VBIDE.Window
Set WinOfBrwObj = FstWinTy(vbext_wt_Browser)
End Property

Property Get BoCompile() As CommandBarButton
Dim O As CommandBarButton
Set O = PopupOfDbg.CommandBar.Controls(1)
If Not HasPfx(O.Caption, "Compi&le") Then Stop
Set BoCompile = O
End Property
Property Get Bars() As Office.CommandBars
Set Bars = BarszV(CVbe)
End Property

Function Bar(BarNm) As Office.CommandBar
Set Bar = Bars(BarNm)
End Function

Private Sub Z()
MIde_CMdMov_Mth:
End Sub
