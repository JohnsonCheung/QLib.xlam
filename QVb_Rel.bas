Attribute VB_Name = "QVb_Rel"
Option Explicit
Private Const CMod$ = "MVb_Rel."
Private Const Asm$ = "QVb"

Property Get SampRel() As Rel
Set SampRel = Rel(SampRelLy)
End Property

Property Get SampRelLy() As String()
Erase XX
X "A B"
X "B A"
SampRelLy = XX
Erase XX
End Property

Property Get SampMthRel() As Rel
Set SampMthRel = Rel(SampMthRelLy)
End Property

Property Get SampMthRelLy() As String()
'SampMthRelLy = RelOf_MthSDNm_To_Mdn_InVbe
End Property


