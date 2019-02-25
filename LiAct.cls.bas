VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiAct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private FxAy() As LiActFx, FbAy() As LiActFb
Friend Function Init(Fx() As LiActFx, Fb() As LiActFb) As LiAct
FxAy = Fx
FbAy = Fb
Set Init = Me
End Function
Property Get Fx() As LiActFx()
Fx = FxAy
End Property
Property Get Fb() As LiActFb()
Fb = FbAy
End Property

