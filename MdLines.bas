VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MdLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const CMod$ = "MdLines."
Public StartLine&, InsLno&, Lines$
Friend Function Init(StartLine, Lines, InsLno) As MdLines
With Me
    .StartLine = StartLine
    .Lines = Lines
    .InsLno = IIf(InsLno = 0, StartLine, InsLno)
End With
Set Init = Me
End Function
Property Get Count&()
Count = LinCnt(Lines)
End Property
