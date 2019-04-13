Attribute VB_Name = "MIde_Cmd_Action"
Option Explicit

Sub TileH()
BtnOfTileV.Execute
End Sub

Sub TileV()
'BtnOfTileV.Execute
End Sub

Property Get TileVBtn() As CommandBarButton
Dim O As CommandBarButton
Set O = PopupOfWin.CommandBar.Controls(3)
If O.Caption <> "Tile &Vertically" Then Stop
Set TileVBtn = O
End Property
