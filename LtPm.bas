VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LtPm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const CMod$ = "LtPm."
Public T$, S$, Cn$
Friend Function Init(T, S, Cn) As LtPm
With Me
    .T = T
    .S = S
    .Cn = Cn
End With
Set Init = Me
End Function

Property Get ToStr$()
ToStr = FmtQQ("T-S-Cn(? ? ?)", T, S, Cn)
End Property
