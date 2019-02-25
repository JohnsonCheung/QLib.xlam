VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lnx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Lin$, Ix&

Friend Sub Init(Lin, Ix&)
Me.Ix = Ix
Me.Lin = Lin
End Sub
Property Get Lno&()
Lno = Ix + 1
End Property
