Attribute VB_Name = "QIde_Gen_Rf_InfDta"
Option Explicit
Private Const CMod$ = "MIde_Gen_Rf_InfDta."
Private Const Asm$ = "QIde"
Private Sub Z_DrsOfRf()
DmpDrs DrsOfRf(CPj)
End Sub
Sub DmpPjRf(P As VBProject)
DmpDrs DrsOfRf(P)
End Sub

Function DrsOfRf(P As VBProject) As Drs
DrsOfRf = Drs(FnyOfRf, DryOfRf(P))
End Function

Property Get FnyOfRf() As String()
FnyOfRf = ItmAddAy("Pj", RfFny)
End Property
Function DryOfRf(P As VBProject) As Variant()
Dim R As vbide.Reference, N$
N = P.Name
For Each R In P.References
    PushI DryOfRf, ItmAddAy(N, DrRf(R))
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

Function DrsOfRfzPjAy(A() As VBProject) As Drs
Dim P
For Each P In Itr(A)
    ApdDrs DrsOfRfzPjAy, DrsOfRf(P)
Next
End Function
