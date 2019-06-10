Attribute VB_Name = "QIde_B_Rf"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Gen_Rf_1."
Private Const Asm$ = "QIde"
Type RfLin: Lin As String: End Type
Public Const DoczFrf$ = "It a file Rf.Txt in Srcp with RfLin"
Public Const FFzRfLin = "Nm Guid Mjr Mnr Frfee"
Function RfLin(Lin) As RfLin
RfLin.Lin = Lin
End Function
Function RfLinzRf(A As VBIDE.Reference) As RfLin
With A
RfLinzRf = RfLin(JnSpcAp(.Name, .GUID, .Major, .Minor, .FullPath))
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
Dim R As VBIDE.Reference
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

Function CvRf(A) As VBIDE.Reference
Set CvRf = A
End Function
Function HasRfNm(Pj As VBProject, RfNm$)
Dim Rf As VBIDE.Reference
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

Function Frfee$(A As VBIDE.Reference)
On Error Resume Next
Frfee = A.FullPath
End Function

Private Sub ZZ()
End Sub

Sub AddRfzDistPj1(DistPj As VBProject)
Stop
'AddRfzRff DistPj, RffzDistPj(DistPj)
End Sub

Sub DmpPjRfP()
DmpDrs DPjRfP
End Sub

Function DPjRfP() As Drs
DPjRfP = DPjRfzP(CPj)
End Function
Function DPjRfzP(P As VBProject) As Drs
Dim FF$: FF = "Name GUID Major Minor FullPath Description BuiltIn Type IsBroken"
Dim A As Drs: A = DrszItrPP(P.References, FF)
DPjRfzP = InsCol(A, "Pj", P.Name)
End Function

Function DPjRfUsr() As Drs
Erase XX
X "MVb"
X "MIde  MVb MXls MAcs"
X "MXls  MVb"
X "MDao  MVb MDta"
X "MAdo  MVb"
X "MAdoX MVb"
X "MApp  MVb"
X "MDta  MVb"
X "MTp   MVb"
X "MSql  MVb"
X "AStkShpCst MVb MXls MAcs"
X "MAcs  MVb MXls"
DPjRfUsr = DrszTRstLy(XX, "Pj Rfnn")
Erase XX
End Function

Function DPjRfStd() As String()
Erase XX
X "MVb   Scripting VBScript_RegExp_55 DAO VBIDE Office"
X "MIde  Scripting VBIDE Excel"
X "MXls  Scripting Office Excel"
X "MDao  Scripting DAO"
X "MAdo  Scripting ADODB"
X "MAdoX Scripting ADOX"
X "MApp  Scripting"
X "MDta  Scripting"
X "MTp   Scripting"
X "MSql  Scripting"
X "AStkShpCst Scripting"
X "MAcs  Scripting Office Access"
DPjRfzStd = DrszTRstLy(XX, "Pj Rfnn")
End Function

Function DStdLib() As Drs
Erase XX
X "VBA                {000204EF-0000-0000-C000-000000000046} 4  2 C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA7.1\VBE7.DLL"
X "Access             {4AFFC9A0-5F99-101B-AF4E-00AA003F0F07} 9  0 C:\Program Files (x86)\Microsoft Office\Root\Office16\MSACC.OLB"
X "stdole             {00020430-0000-0000-C000-000000000046} 2  0 C:\Windows\SysWOW64\stdole2.tlb"
X "Excel              {00020813-0000-0000-C000-000000000046} 1  9 C:\Program Files (x86)\Microsoft Office\Root\Office16\EXCEL.EXE"
X "Scripting          {420B2830-E718-11CF-893D-00A0C9054228} 1  0 C:\Windows\SysWOW64\scrrun.dll"
X "DAO                {4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28} 12 0 C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\ACEDAO.DLL"
X "Office             {2DF8D04C-5BFA-101B-BDE5-00AA0044DE52} 2  8 C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\MSO.DLL"
X "ADODB              {B691E011-1797-432E-907A-4D8C69339129} 6  1 C:\Program Files (x86)\Common Files\System\ado\msado15.dll"
X "ADOX               {00000600-0000-0010-8000-00AA006D2EA4} 6  0 C:\Program Files (x86)\Common Files\System\ado\msadox.dll"
X "VBScript_RegExp_55 {3F4DACA7-160D-11D2-A8E9-00104B365C9F} 5  5 C:\Windows\SysWOW64\vbscript.dll"
X "VBIDE              {0002E157-0000-0000-C000-000000000046} 5  3 C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
DStdLib = Drsz4TRstLy(XX, "Libn Guid Maj Mnr Ffn")
Erase XX
End Function

Private Sub Z_FAny_DPD_ORD()
'GoSub ZZ
GoSub T1
Exit Sub
T1:
    Ept = SyzSS("MVb MXls MAdo MAdoX MApp MDta MTp MSql MDao MAcs MIde AStkShpCst")
    GoTo Tst
Tst:
'    Act = FAny_DPD_ORD
    C
    Return
ZZ:
    ClrImm
    D "Rel --------------------"
    D UsrPjRfLy
    D "Itms-DPD-ORD --------------------"
'   D FAny_DPD_ORD
    Return
End Sub

Sub BrwPjRfvStd()
Brw FmtSyT4(RfSrczSrcp(SrcpP))
End Sub
Sub DmpRf()
DmpRfzP CPj
End Sub
Sub DmpRfzP(P As VBProject)
D FmtSyT4(RfSrc(P))
End Sub

Private Function GuidLinRfNm$(RfNm)
Dim D As Dictionary
'Set D = RfDfn_STD
If D.Exists(RfNm) Then GuidLinRfNm = D(RfNm): Exit Function
Thw CSub, "Given RfNm cannot find the STD GUID Dic", "RfNm RfDfn_STD", RfNm, FmtDic(D)
End Function

Private Function GuidLinyPjn(Pjn$) As String()
Dim RfNm
For Each RfNm In Itr(StdRfNyPj(Pjn))
    PushI GuidLinyPjn, RfNm & " " & GuidLinRfNm(RfNm)
Next
End Function

Private Function StdRfNyPj(Pjn$) As String()
Dim D As Dictionary
Set D = PjnToStdRfNNDic
If Not D.Exists(Pjn) Then
    Thw CSub, "Given Pj not found in RfDfn_STD", "Pj RfDfn_STD", Pjn, FmtDic(PjnToStdRfNNDic)
End If
StdRfNyPj = SyzSS(D(Pjn))
End Function



