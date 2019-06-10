VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim WithEvents BtnOfRunZ As CommandBarButton
Attribute BtnOfRunZ.VB_VarHelpID = -1
Dim WithEvents BtnOfAlignMth As CommandBarButton
Attribute BtnOfAlignMth.VB_VarHelpID = -1

Private Property Get Y_BtnSpec() As String()
Erase XX
X "Bars"
X " XX AlignMth RunZ"
Y_BtnSpec = XX
Erase XX
End Property

Private Sub BtnOfAlignMth_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
AlignMthDim
End Sub

Private Sub BtnOfRunZ_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
Z
End Sub

Friend Sub Class_Initialize()
Dim Spec$():              Spec = Y_BtnSpec  ' Spec
EnsBtns IndentedLy(Spec, "Bars")
Set BtnOfRunZ = Bars("XX").Controls("RunZ")
Set BtnOfAlignMth = Bars("XX").Controls("AlignMth")
'BtnOfRunZ.ShortcutText = "Alt+Z"
End Sub
Private Sub Class_Terminate()
MsgBox "Class1 terminated"
End Sub
