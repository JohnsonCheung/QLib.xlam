Attribute VB_Name = "QIde_Tool_IdeTool"
Option Explicit
Option Compare Text
Function AAA(X)
Dim Rg As Range
Set Rg = CWs.Range("D5")
Rg.Value = 123
AAA = X + 1
Exit Function
X: Debug.Print Err.Description
End Function
Private Function Spec() As String()
Erase XX
X "Bars"
X " AA A1 A2 A3"
X " BB B1 B2 B3"
X "Btns"
X " A1"
Spec = XX  '*Spec
Erase XX
End Function

Private Function BtnSpec() As String()
BtnSpec = IndentedLy(Spec, "Bars")
End Function

Sub InstallIdeTools()
EnsBtns BtnSpec
EnsClsLines "IdeTool", ToolClsCd
End Sub

Private Function ToolClsCd$()
Stop
End Function

Private Function ToolBarNy() As String()
ToolBarNy = T1Ay(BtnSpec)
End Function

Sub RmvIdeTools()
RmvBarNy ToolBarNy
End Sub
