VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SyPair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private X_Sy1$(), X_Sy2$()
Function Init(Sy1, Sy2) As SyPair
X_Sy1 = Sy1
X_Sy2 = Sy2
Set Init = Me
End Function
Property Get Sy1() As String()
Sy1 = X_Sy1
End Property
Property Get Sy2() As String()
Sy2 = X_Sy2
End Property

