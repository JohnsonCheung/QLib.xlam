Attribute VB_Name = "QIde_Ctl_Ctl"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Btn."
Private Const Asm$ = "QIde"

Sub DltClr(A As CommandBar)
Dim I
For Each I In Itr(CtlAy(A))
    CvCtl(I).Delete
Next
End Sub

Function CtlAy(A As CommandBar) As CommandBarControl()
CtlAy = IntozItr(CtlAy, A.Controls)
End Function

Function CtlCapNy(A As CommandBar) As String()
CtlCapNy = SyzItrP(CtlAy(A), PrpPth("Caption"))
End Function

Property Get BarNy() As String()
BarNy = BarNyzV(CVbe)
End Property

Property Get WinOfBrwObj() As vbIde.Window
Set WinOfBrwObj = FstWinTy(vbext_wt_Browser)
End Property

Property Get BtnOfCompile() As CommandBarButton
Dim O As CommandBarButton
Set O = PopupOfDbg.CommandBar.Controls(1)
If Not HasPfx(O.Caption, "Compi&le") Then Stop
Set BtnOfCompile = O
End Property

Private Function CvCtl(A) As CommandBarControl
Set CvCtl = A
End Function

Property Get PopupOfDbg() As CommandBarPopup
Set PopupOfDbg = BarOfMnu.Controls("Debug")
End Property

Property Get BtnOfEdtClr() As Office.CommandBarButton
Set BtnOfEdtClr = FstCaption(EditPopup.Controls, "C&lear")
End Property

Property Get BarOfMnu() As CommandBar
Set BarOfMnu = BarOfMnuzV(CVbe)
End Property

Property Get BtnOfSelAll() As Office.CommandBarButton
Set BtnOfSelAll = FstCaption(EditPopup.Controls, "Select &All")
End Property

Function IsBtn(A) As Boolean
IsBtn = TypeName(A) = "CommandButton"
End Function

Property Get BtnOfNxtStmt() As CommandBarButton
Set BtnOfNxtStmt = PopupOfDbg.Controls("Show Next Statement")
End Property
Private Function FstCaption(Itr, Caption$) 'Return FstItm with Caption-Prp = Caption$
FstCaption = FstItmPEv(Itr, PrpPth("Caption"), Caption)
End Function
Private Property Get EditPopup() As Office.CommandBarPopup
Set EditPopup = FstCaption(BarOfMnu.Controls, "&Edit")
End Property

Property Get BtnOfSav() As CommandBarButton
Set BtnOfSav = BtnOfSavzV(CVbe)
End Property

Property Get BtnOfJmpNxtStmt() As CommandBarButton
Set BtnOfJmpNxtStmt = PopupOfDbg.Controls("Show Next Statement")
End Property

Property Get StdBar() As Office.CommandBar
Set StdBar = Bars("Standard")
End Property
Function BarOfMnuzV(A As Vbe) As CommandBar
Set BarOfMnuzV = A.CommandBars("Menu Bar")
End Function

Function BtnOfSavzV(A As Vbe) As CommandBarButton
Dim I As CommandBarControl, S As Office.CommandBarControls
Set S = StdBarzV(A).Controls
For Each I In S
    If HasPfx(I.Caption, "&Sav") Then Set BtnOfSav = I: Exit Function
Next
Stop
End Function

Function StdBarzV(A As Vbe) As Office.CommandBar
Dim X As Office.CommandBars
Set X = BarszV(A)
Set StdBarzV = X("Standard")
End Function

Function BarAyzV(A As Vbe) As Office.CommandBar()
Dim I
For Each I In A.CommandBars
   PushObj BarAyzV, I
Next
End Function

Function BarNyzV(A As Vbe) As String()
BarNyzV = Itn(A.CommandBars)
End Function

Property Get PopupOfWin() As CommandBarPopup
Set PopupOfWin = BarOfMnu.Controls("Window")
End Property
Property Get BtnOfTileH() As CommandBarButton
Set BtnOfTileH = PopupOfWin.Controls("Tile &Horizontally")
End Property

Property Get BtnOfTileV() As Office.CommandBarButton
Set BtnOfTileV = PopupOfWin.Controls("Tile &Vertically")
End Property

Property Get BtnOfXls() As Office.CommandBarControl
Set BtnOfXls = StdBar.Controls(1)
End Property

Private Sub ZZ_DbgPop()
Dim A
Set A = PopupOfDbg
Stop
End Sub

Private Sub ZZ_BarOfMnu()
Dim A As CommandBar
Set A = BarOfMnu
Stop
End Sub

Private Sub ZZ()
Dim A As CommandBar
Dim B As Variant
Dim C As Vbe
DltClr A
CtlAy A
IsBtn B
BarNyzV C
BarAyzV C
BarNyzV C
End Sub

