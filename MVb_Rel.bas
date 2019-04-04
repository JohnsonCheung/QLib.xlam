Attribute VB_Name = "MVb_Rel"
Option Explicit
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
SampMthRelLy = RelOf_MthSDNm_To_MdNm_OfVbe
End Property


