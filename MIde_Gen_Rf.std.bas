Attribute VB_Name = "MIde_Gen_Rf"
Option Explicit
Const CMod$ = "MIde_Pj_Rf."

Property Get RffAy() As String()
RffAy = RffAyPj(CurPj)
End Property

Property Get FmtRf() As String()
FmtRf = AyAlign3T(RfLy)
End Property

Property Get RfLy() As String()
RfLy = RfSrc(CurPj)
End Property

Function RfNyPj(A As VBProject) As String()
RfNyPj = Itn(A.References)
End Function
Property Get RfNy() As String()
RfNy = RfNyPj(CurPj)
End Property

Function CvRf(A) As VBIDE.Reference
Set CvRf = A
End Function

Sub AddRfzAy(Pj As VBProject, RffAy$())

End Sub

Sub CpyPjRfToPj(Pj As VBProject, ToPj As VBProject)
AddRfzAy ToPj, RffAyPj(Pj)
End Sub

Function HasRf(Pj As VBProject, RfNm)
Dim Rf As VBIDE.Reference
For Each Rf In Pj.References
    If Rf.Name = RfNm Then HasRf = True: Exit Function
Next
End Function
Function HasRfGuid(A As VBProject, RfGuid)
HasRfGuid = HasItrPEv(A.References, "GUID", RfGuid)
End Function

Function HasRff(A As VBProject, Rff) As Boolean
HasRff = HasItrPEv(A.References, "FullPath", Rff)
End Function

Sub BrwRf()
BrwAy FmtRf
End Sub

Function RffAyPj(A As VBProject) As String()
RffAyPj = SyItrPrp(A.References, "FullPath")
End Function

Function RfLin$(A As VBIDE.Reference)
With A
RfLin = .Name & " " & .Guid & " " & .Major & " " & .Minor & " " & .FullPath
End With
End Function


Function RffPjNm$(A As VBProject, RfNm$)
RffPjNm = PthPj(A) & RfNm & ".xlam"
End Function

Function PjRfNy(A As VBProject) As String()
PjRfNy = Itn(A.References)
End Function

Sub RmvPjRfNN(A As VBProject, RfNN$)
Dim N
For Each N In Ny(RfNN)
    'RmvPjRf A, N
Next
SavPj A
End Sub
Function StdRff$(StdRfNm)

End Function

Sub AddPjStdRf(A As VBProject, StdRfNm)
Const CSub$ = CMod & "AddPjStdRf"
If HasRf(A, StdRfNm) Then
    Debug.Print FmtQQ("AddPjStdRf: Pj(?) already has StdRfNm(?)", A.Name, StdRfNm)
    Exit Sub
End If
Dim Rff$: Rff = StdRff(StdRfNm)
ThwNotExistFfn Rff, CSub, "StdRfFil"
A.References.AddFromFile Rff
End Sub

Function Rff$(A As VBIDE.Reference)
On Error Resume Next
Rff = A.FullPath
End Function

Function RfPth$(A As VBIDE.Reference)
On Error Resume Next
RfPth = A.FullPath
End Function

Function RfToStr$(A As VBIDE.Reference)
With A
   RfToStr = .Name & " " & RfPth(A)
End With
End Function
Private Sub ZZ()
Dim A As Variant
Dim B As VBProject
Dim C$()
Dim D$
Dim E As Reference
Dim F As VBIDE.Reference
CvRf A
AddRfzAy B, C
HasRf B, D
PjRfNy B
RmvPjRfNN B, D
End Sub

Private Sub Z()
End Sub
