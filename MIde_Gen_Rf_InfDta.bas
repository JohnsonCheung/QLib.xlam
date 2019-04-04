Attribute VB_Name = "MIde_Gen_Rf_InfDta"
Option Explicit
Private Sub Z_PjRfDrs()
DmpDrs PjRfDrs(CurPj)
End Sub
Sub DmpPjRf(A As VBProject)
DmpDrs PjRfDrs(A)
End Sub
Function PjRfDrs(A As VBProject) As Drs
Set PjRfDrs = Drs(PjRfFny, PjRfDry(A))
End Function
Property Get PjRfFny() As String()
PjRfFny = AyItmAddAy("Pj", RfFny)
End Property
Function PjRfDry(A As VBProject) As Variant()
Dim R As Vbide.Reference, N$
N = A.Name
For Each R In A.References
    PushI PjRfDry, AyItmAddAy(N, DrRf(R))
Next
End Function
Function DrRf(A As Vbide.Reference) As Variant()
With A
DrRf = Array(.Name, .Guid, .Major, .Minor, .FullPath, .Description, .BuiltIn, .Type, .IsBroken)
End With
End Function
Property Get RfFny() As String()
RfFny = SySsl(RmvDotComma(".Name, .GUID, .Major, .Minor, .FullPath, .Description, .BuiltIn, .Type, .IsBroken"))
End Property

Function PjAyRfDrs(A() As VBProject) As Drs
Dim P
For Each P In Itr(A)
    PushDrs PjAyRfDrs, PjRfDrs(CvPj(P))
Next
End Function
