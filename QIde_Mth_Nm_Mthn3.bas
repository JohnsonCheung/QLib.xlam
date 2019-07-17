Attribute VB_Name = "QIde_Mth_Nm_Mthn3"
Option Compare Text
Option Explicit
Private Const CMod$ = "Mthn3."
Type Mthn3: Nm As String: ShtTy As String: ShtMdy As String: End Type
':Mthn_ZDash$ = "Mthn rule.  When beg with Z_, it is trying to test.  Try use own resource, like Y_.  Don't use other md resource."
':Mthn_ZZDash$ = "Mthn rule.  When beg with Z_, it is tested ok.  It should always pass.  Using Z_ due to it sinks to bottom."
':Mthn_Z$ = "Mthn rule.  A private mth with all Z_ fun and a Lbl eq the mdn."
':Mthn_YDash$ = "Mthn rule.  PurePrp for testing.  Used to Z_ mth."

Function Mthn3(Nm, ShtMdy, ShtTy) As Mthn3
With Mthn3
    .Nm = Nm
    .ShtMdy = ShtMdy
    .ShtTy = ShtTy
End With
End Function

Function Mthn3zL(Lin) As Mthn3
Mthn3zL = ShfMthn3(CStr(Lin))
End Function


Function ShfMthn3(OLin$) As Mthn3
Dim M$: M = ShfShtMdy(OLin)
Dim T$: T = ShfShtMthTy(OLin):: If T = "" Then Exit Function
ShfMthn3 = Mthn3(ShfNm(OLin), M, T)
End Function

Function RmvMthn3$(Lin)
Dim L$: L = Lin
RmvMthMdy L
If ShfMthTy(L) = "" Then Exit Function
If ShfNm(L) = "" Then Thw CSub, "Not as SrcLin", "Lin", Lin
RmvMthn3 = L
End Function

Function FmtMthn3$(A As Mthn3)
With A
FmtMthn3 = JnDotAp(.Nm, .ShtMdy, .ShtTy)
End With
End Function
Sub DmPubMthn3(A As Mthn3)
D FmtMthn3(A)
End Sub


