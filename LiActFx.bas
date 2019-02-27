VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiActFx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Fx As String
Public Fxn As String
Public Wsn As String
Public ShtTyDic As Dictionary
Friend Function Init(Fx, Fxn, Wsn, ShtTyDic As Dictionary) As LiActFx
With Me
    .Fx = Fx
    Set .ShtTyDic = ShtTyDic
    .Fxn = Fxn
    .Wsn = Wsn
End With
Set Init = Me
End Function

Property Get Fset() As Aset
Set Fset = AsetzDicKey(ShtTyDic)
End Property

