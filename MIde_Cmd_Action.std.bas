Attribute VB_Name = "MIde_Cmd_Action"
Option Explicit

Sub TileH()
WinTileVertBtn.Execute
End Sub

Sub TileV()
WinTileVertBtn.Execute
End Sub

Property Get TileVBtn() As CommandBarButton
Dim O As CommandBarButton
Set O = WinPop.CommandBar.Controls(3)
If O.Caption <> "Tile &Vertically" Then Stop
Set TileVBtn = O
End Property
