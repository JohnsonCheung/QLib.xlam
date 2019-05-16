Attribute VB_Name = "QIde_Gen_Rf_InfDta"
Option Explicit
Private Const CMod$ = "MIde_Gen_Rf_InfDta."
Private Const Asm$ = "QIde"
Private Sub Z_Drs_Rf()
DmpDrs Drs_Rf(CPj)
End Sub
Sub DmpPjRf(P As VBProject)
DmpDrs Drs_Rf(P)
End Sub

Function Drs_Rf(P As VBProject) As Drs
Drs_Rf = Drs(Fny_Rf, Dry_Rf(P))
End Function

Property Get Fny_Rf() As String()
Fny_Rf = ItmAddAy("Pj", RfFny)
End Property
Function Dry_Rf(P As VBProject) As Variant()
Dim R As vbide.Reference, N$
N = P.Name
For Each R In P.References
    PushI Dry_Rf, ItmAddAy(N, DrRf(R))
Next
End Function
Function DrRf(A As vbide.Reference) As Variant()
With A
DrRf = Array(.Name, .Guid, .Major, .Minor, .FullPath, .Description, .BuiltIn, .Type, .IsBroken)
End With
End Function
Property Get RfFny() As String()
RfFny = SyzSS(RmvDotComma(".Name, .GUID, .Major, .Minor, .FullPath, .Description, .BuiltIn, .Type, .IsBroken"))
End Property

Function Drs_RfzPjAy(A() As VBProject) As Drs
Dim P
For Each P In Itr(A)
    ApdDrs Drs_RfzPjAy, Drs_Rf(CvPj(P))
Next
End Function
