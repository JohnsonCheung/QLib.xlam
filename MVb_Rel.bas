Attribute VB_Name = "MVb_Rel"
Option Explicit
Property Get SampRel() As Rel
Set SampRel = Rel(SampRelLy)
End Property
Property Get SampRelLy() As String()
Erase xx
X "A B"
X "B A"
SampRelLy = xx
Erase xx
End Property

Property Get SampMthRel() As Rel
Set SampMthRel = Rel(SampMthRelLy)
End Property

Property Get SampMthRelLy() As String()
SampMthRelLy = RelOf_MthSDNm_To_MdNm_InVbe
End Property


