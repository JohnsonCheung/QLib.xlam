VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "Class1."
Dim WithEvents BoRunZ As CommandBarButton
Attribute BoRunZ.VB_VarHelpID = -1
Dim WithEvents BoAlignMth As CommandBarButton
Attribute BoAlignMth.VB_VarHelpID = -1

Property Get Y_BtnSpec() As String()
Erase XX
X "Bars"
X " XX AlignMth RunZ"
Y_BtnSpec = XX
Erase XX
End Property

Sub BoAlignMth_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
AlignMth
End Sub

Friend Sub Class_Initialize()
Dim Spec$():              Spec = Y_BtnSpec  ' Spec
EnsBtns IndentedLy(Spec, "Bars")
Set BoRunZ = Bars("XX").Controls("RunZ")
Set BoAlignMth = Bars("XX").Controls("AlignMth")
'BoRunZ.ShortcutText = "Alt+Z"
End Sub

Sub Class_Terminate()
MsgBox "Class1 terminated"
End Sub
