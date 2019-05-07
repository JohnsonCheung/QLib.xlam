Attribute VB_Name = "QIde_Btn"
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

Function CtlCapAy(A As CommandBar) As String()
CtlCapAy = SyzItrP(CtlAy(A), "Caption")
End Function
Function Bar(Nm$) As CommandBar
Set Bar = CurVbe.CommandBars(Nm)
End Function
Property Get BarNy() As String()
BarNy = BarNyByVbe(CurVbe)
End Property

Property Get BrwObjWin() As VBIDE.Window
Set BrwObjWin = FstWinTy(vbext_wt_Browser)
End Property

Property Get CompileBtn() As CommandBarButton
Dim O As CommandBarButton
Set O = DbgPopup.CommandBar.Controls(1)
If Not HasPfx(O.Caption, "Compi&le") Then Stop
Set CompileBtn = O
End Property

Private Function CvCtl(A) As CommandBarControl
Set CvCtl = A
End Function

Property Get DbgPopup() As CommandBarPopup
Set DbgPopup = MnuBar.Controls("Debug")
End Property

Property Get EdtclrBtn() As Office.CommandBarButton
Set EdtclrBtn = FstItmPEv(EditPopup.Controls, "Caption", "C&lear")
End Property

Property Get MnuBar() As CommandBar
Set MnuBar = MnuBarzVbe(CurVbe)
End Property

Property Get SelAllBtn() As Office.CommandBarButton
Set SelAllBtn = FstItmPEv(EditPopup.Controls, "Caption", "Select &All")
End Property

Function IsBtn(A) As Boolean
IsBtn = TypeName(A) = "CommandButton"
End Function


Property Get NxtStmtBtn() As CommandBarButton
Set NxtStmtBtn = DbgPopup.Controls("Show Next Statement")
End Property

Private Property Get EditPopup() As Office.CommandBarPopup
Set EditPopup = FstItmPEv(MnuBar.Controls, "Caption", "&Edit")
End Property

Property Get SavBtn() As CommandBarButton
Set SavBtn = SavBtnzVbe(CurVbe)
End Property

Property Get JmpNxtStmtBtn() As CommandBarButton
Set JmpNxtStmtBtn = DbgPopup.Controls("Show Next Statement")
End Property

Property Get StdBar() As Office.CommandBar
Set StdBar = Bars("Standard")
End Property
Function MnuBarzVbe(A As Vbe) As CommandBar
Set MnuBarzVbe = A.CommandBars("Menu Bar")
End Function

Function SavBtnzVbe(A As Vbe) As CommandBarButton
Dim I As CommandBarControl, S As Office.CommandBarControls
Set S = StdBarzVbe(A).Controls
For Each I In S
    If HasPfx(I.Caption, "&Sav") Then Set SavBtn = I: Exit Function
Next
Stop
End Function

Function StdBarzVbe(A As Vbe) As Office.CommandBar
Dim X As Office.CommandBars
Set X = BarszVbe(A)
Set StdBarzVbe = X("Standard")
End Function

Function BarNyByVbe(A As Vbe) As String()
BarNyByVbe = CmdBarNyvVbe(A)
End Function

Function BarAyByVbe(A As Vbe) As Office.CommandBar()
Dim I
For Each I In A.CommandBars
   PushObj BarAyByVbe, I
Next
End Function

Function CmdBarNyvVbe(A As Vbe) As String()
CmdBarNyvVbe = Itn(A.CommandBars)
End Function

Property Get PopupOfWin() As CommandBarPopup
Set PopupOfWin = MnuBar.Controls("Window")
End Property

Property Get BtnOfTileV() As Office.CommandBarButton
Set BtnOfTileV = PopupOfWin.Controls("Tile &Vertically")
End Property

Property Get BtnOfXls() As Office.CommandBarControl
Set BtnOfXls = StdBar.Controls(1)
End Property

Private Sub ZZ_DbgPop()
Dim A
Set A = DbgPopup
Stop
End Sub

Private Sub ZZ_MnuBar()
Dim A As CommandBar
Set A = MnuBar
Stop
End Sub

Private Sub ZZ()
Dim A As CommandBar
Dim B As Variant
Dim C As Vbe
DltClr A
CtlAy A
IsBtn B
BarNyByVbe C
BarAyByVbe C
CmdBarNyvVbe C
End Sub

Private Sub Z()
End Sub
