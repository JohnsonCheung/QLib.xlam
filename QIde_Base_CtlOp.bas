Attribute VB_Name = "QIde_Base_CtlOp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Cmd_Action."
Private Const Asm$ = "QIde"
Sub TileH()
BtnOfTileH.Execute
End Sub
Sub TileV()
BtnOfTileV.Execute
End Sub
Sub Compile(Pjn$)
JmpzP Pj(Pjn)
BtnOfCompile.Execute
End Sub
Sub CompilezP(P As VBProject)
JmpzP P
ThwIf_BtnOfCompile P.Name
With BtnOfCompile
    If .Enabled Then
        .Execute
        Debug.Print P.Name, "<--- Compiled"
    Else
        Debug.Print P.Name, "already Compiled"
    End If
End With
BtnOfTileV.Execute
BtnOfSav.Execute
End Sub

Sub CompilezV(A As Vbe)
DoItrFun A.VBProjects, "CompilezP"
End Sub

Sub ThwIf_BtnOfCompile(NEPjn$)
Dim Act$, Ept$
Act = BtnOfCompile.Caption
Ept = "Compi&le " & NEPjn
If Act <> Ept Then Thw CSub, "Cur BtnOfCompile.Caption <> Compi&le {Pjn}", "Compile-Btn-Caption Pjn Ept-Btn-Caption", Act, NEPjn, Ept
End Sub

Private Sub Z_PjCompile()
CompilezP CPj
End Sub

Sub DltClr(A As CommandBar)
Dim I
For Each I In Itr(OyzItr(A.Controls))
    CvCtl(I).Delete
Next
End Sub

Private Sub ZZ()
Dim A As CommandBar
Dim B As Variant
Dim C As Vbe
DltClr A
IsBtn B
BarNyzV C
BarNyzV C
End Sub
Sub DltBar(BarNm$)
If Not HasBar(BarNm) Then Debug.Print "Bar[" & BarNm & "] not found": Exit Sub
Bars(BarNm).Delete
End Sub
Private Sub ZZ_EnsBtns()
Dim Spec$(), BarBtnccAy$()
Erase XX
X "Bars"
X " AA A1 A2 A3"
X " BB B1 B2 B3"
X "Btns"
X " A1"
Spec = XX  '*Spec
Erase XX
BarBtnccAy = IndentedLy(Spec, "Bars")
EnsBtns BarBtnccAy
End Sub
Sub EnsBtns(BarBtnccAy$())
Dim I
For Each I In Itr(BarBtnccAy)
    EnsBarBtncc I
Next
End Sub

Private Sub EnsBarBtncc(BarBtncc)
Dim L$
L = BarBtncc
EnsBtnzCC EnsBar(ShfT1(L)), L
End Sub
Sub RmvBarNy(BarNy$())
Dim IBar
For Each IBar In BarNy
    If HasBar(IBar) Then
        If Not Bar(IBar).BuiltIn Then
            Bar(IBar).Delete
        End If
    End If
Next
End Sub
Private Function EnsBar(BarNm$) As CommandBar
If HasBar(BarNm) Then
    Set EnsBar = Bars(BarNm)
Else
    Set EnsBar = Bars.Add(BarNm)
End If
EnsBar.Visible = True
End Function
Private Sub EnsBtnzCC(Bar As CommandBar, BtnCapcc$)
Dim BtnCap
For Each BtnCap In TermAy(BtnCapcc)
    EnsBtnzC Bar, BtnCap
Next
End Sub
Private Function HasBtn(Bar As CommandBar, BtnCap) As Boolean
Dim C As CommandBarControl
For Each C In Bar.Controls
    If C.Type = msoControlButton Then
        If C.Caption = BtnCap Then HasBtn = True: Exit Function
    End If
Next

End Function

Private Sub EnsBtnzC(Bar As CommandBar, BtnCap)
If HasBtn(Bar, BtnCap) Then Exit Sub
Dim B As CommandBarButton
Set B = Bar.Controls.Add(MsoControlType.msoControlButton)
B.Caption = BtnCap
B.Style = msoButtonCaption
End Sub

Private Sub AddBtn(Bar As CommandBar, BtnCap)
Dim B As CommandBarButton
Set B = Bar.Controls.Add(MsoControlType.msoControlButton)
B.Caption = BtnCap
B.Style = msoButtonCaption
End Sub

Private Sub ABtn_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
Stop
End Sub

Sub ZZZ()
QIde_Base_CtlOp:
End Sub

