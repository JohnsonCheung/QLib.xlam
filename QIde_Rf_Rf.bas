Attribute VB_Name = "QIde_Rf_Rf"
Option Explicit
Const Asm$ = "MIde"
Const Ns$ = "Ide.PjInf"
Private Const CMod$ = "BRf."
Type RfLin: Lin As String: End Type
Public Const DoczFrf$ = "It a file Rf.Txt in Srcp with RfLin"
Public Const FFzRfLin$ = "Nm Guid Mjr Mnr Frfee"
Function RfLin(Lin$) As RfLin
RfLin.Lin = Lin
End Function
Function RfLinzRf(A As VBIDE.Reference) As RfLin
With A
RfLinzRf = RfLin(JnSpcAp(.Name, .Guid, .Major, .Minor, .FullPath))
End With
End Function
Function FrfeezRfLin$(A As RfLin)
Dim P%: P = InStr(Replace(A.Lin, " ", "-", Count:=3), " ")
FrfeezRfLin = Mid(A.Lin, P + 1)
End Function

Function HasFrfee(A As VBProject, Frfee$) As Boolean
HasFrfee = HasItrPEv(A.References, "FullPath", Frfee)
End Function
Property Get FrfC$()
FrfC = Frf(CurPj)
End Property

Function FrfzSrcp$(Srcp$)
FrfzSrcp = EnsPthSfx(Srcp) & "Rf.txt"
End Function

Function Frf$(A As VBProject)
Frf = FrfzSrcp(Srcp(A))
End Function

Function FrfzDistPj$(DistPj As VBProject)
FrfzDistPj = SrcpzDistPj(DistPj) & "Rf.txt"
End Function

Function RfSrcC() As String()
RfSrcC = RfSrc(CurPj)
End Function

Function RfSrczSrcp(Srcp$) As String()
RfSrczSrcp = LyzFt(FrfzSrcp(Srcp))
End Function
Function RfSrc(A As VBProject) As String()
Dim R As VBIDE.Reference
For Each R In A.References
    PushI RfSrc, RfLinzRf(R).Lin
Next
End Function

Property Get RffSy() As String()
RffSy = RffSyzPj(CurPj)
End Property

Property Get FmtRfPj() As String()
FmtRfPj = FmtSyT3(RfLyPj)
End Property

Property Get RfLyPj() As String()
RfLyPj = RfSrczPj(CurPj)
End Property

Function RfNyzPj(A As VBProject) As String()
RfNyzPj = Itn(A.References)
End Function

Property Get RfNyP() As String()
RfNy = RfNyPj(CurPj)
End Property

Function CvRf(A) As VBIDE.Reference
Set CvRf = A
End Function
Function HasRfNm(Pj As VBProject, RfNm$)
Dim Rf As VBIDE.Reference
For Each Rf In Pj.References
    If Rf.Name = RfNm Then HasRfNm = True: Exit Function
Next
End Function

Function HasRfGuid(A As VBProject, RfGuid)
HasRfGuid = HasItrPEv(A.References, "GUID", RfGuid)
End Function

Sub BrwRf()
BrwAy FmtRfPj
End Sub

Function RffSyzPj(A As VBProject) As String()
RffSyzPj = SyzItrPrp(A.References, "FullPath")
End Function

Function RffPjNm$(A As VBProject, RfNm$)
RffPjNm = PjPth(A) & RfNm & ".xlam"
End Function

Function PjRfNy(A As VBProject) As String()
PjRfNy = Itn(A.References)
End Function

Function Frfee$(A As VBIDE.Reference)
On Error Resume Next
Frfee = A.FullPath
End Function

Private Sub Z()
End Sub

Sub AddRfzDistPj1(DistPj As VBProject)
AddRfzRff DistPj, RffzDistPj(DistPj)
End Sub

