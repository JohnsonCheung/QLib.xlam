Attribute VB_Name = "QVb_Dta_Rel"
Option Compare Text
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

Property Get SamPubMthRel() As Rel
Set SamPubMthRel = Rel(SamPubMthRelLy)
End Property

Property Get SamPubMthRelLy() As String()
'SampMthRelLy = RelOf_MthSDNm_To_Mdn_InVbe
End Property


