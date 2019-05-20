Attribute VB_Name = "QIde_Rf_Rf"
Option Compare Text
Option Explicit
Const Asm$ = "MIde"
Const NS$ = "Ide.PjInf"
Private Const CMod$ = "BRf."
Type RfLin: Lin As String: End Type
Public Const DoczFrf$ = "It a file Rf.Txt in Srcp with RfLin"
Public Const FFzRfLin = "Nm Guid Mjr Mnr Frfee"
Function RfLin(Lin) As RfLin
RfLin.Lin = Lin
End Function
Function RfLinzRf(A As vbIde.Reference) As RfLin
With A
RfLinzRf = RfLin(JnSpcAp(.Name, .Guid, .Major, .Minor, .FullPath))
End With
End Function
Function FrfeezRfLin(A As RfLin)
Dim P%: P = InStr(Replace(A.Lin, " ", "-", Count:=3), " ")
FrfeezRfLin = Mid(A.Lin, P + 1)
End Function

Function HasFrfee(P As VBProject, Frfee) As Boolean
HasFrfee = HasItrPEv(P.References, "FullPath", Frfee)
End Function
Property Get FrfC$()
FrfC = Frf(CPj)
End Property

Function FrfzSrcp$(Srcp$)
FrfzSrcp = EnsPthSfx(Srcp) & "Rf.txt"
End Function

Function Frf$(P As VBProject)
Frf = FrfzSrcp(Srcp(P))
End Function

Function FrfzDistPj$(DistPj As VBProject)
FrfzDistPj = SrcpzDistPj(DistPj) & "Rf.txt"
End Function

Function RfSrcC() As String()
RfSrcC = RfSrc(CPj)
End Function

Function RfSrczSrcp(Srcp$) As String()
RfSrczSrcp = LyzFt(FrfzSrcp(Srcp))
End Function
Function RfSrc(P As VBProject) As String()
Dim R As vbIde.Reference
For Each R In P.References
    PushI RfSrc, RfLinzRf(R).Lin
Next
End Function

Property Get RffSy() As String()
RffSy = RffyzP(CPj)
End Property

Property Get FmtRfPj() As String()
FmtRfPj = FmtSy3Term(RfLyPj)
End Property

Property Get RfLyPj() As String()
Stop
'RfLyPj = RfSrczP(CPj)
End Property

Function RfNyzP(P As VBProject) As String()
RfNyzP = Itn(P.References)
End Function

Property Get RfNyP() As String()
RfNyP = RfNyzP(CPj)
End Property

Function CvRf(A) As vbIde.Reference
Set CvRf = A
End Function
Function HasRfNm(Pj As VBProject, RfNm$)
Dim Rf As vbIde.Reference
For Each Rf In Pj.References
    If Rf.Name = RfNm Then HasRfNm = True: Exit Function
Next
End Function

Function HasRfGuid(P As VBProject, RfGuid)
HasRfGuid = HasItrPEv(P.References, "GUID", RfGuid)
End Function

Sub BrwRf()
BrwAy FmtRfPj
End Sub

Function RffyzP(P As VBProject) As String()
RffyzP = SyzItrPrp(P.References, "FullPath")
End Function

Function RffPjn$(P As VBProject, RfNm$)
RffPjn = Pjp(P) & RfNm & ".xlam"
End Function

Function PjRfNy(P As VBProject) As String()
PjRfNy = Itn(P.References)
End Function

Function Frfee$(A As vbIde.Reference)
On Error Resume Next
Frfee = A.FullPath
End Function

Private Sub ZZ()
End Sub

Sub AddRfzDistPj1(DistPj As VBProject)
Stop
'AddRfzRff DistPj, RffzDistPj(DistPj)
End Sub

