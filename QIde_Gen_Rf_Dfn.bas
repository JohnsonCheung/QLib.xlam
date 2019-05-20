Attribute VB_Name = "QIde_Gen_Rf_Dfn"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Gen_Rf_Dfn."
Private Const Asm$ = "QIde"
Private Property Get UsrPjRfLy() As String()
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
UsrPjRfLy = XX
Erase XX
End Property

Property Get PjnToStdRfNNDic() As Dictionary
Static O As Dictionary
If IsNothing(O) Then Set O = Dic(PjnToStdRfNNLy)
Set PjnToStdRfNNDic = O
End Property
Private Property Get PjnToStdRfNNLy() As String()
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
PjnToStdRfNNLy = XX
Erase XX
End Property

Private Property Get StdGuidLy() As String()
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
StdGuidLy = XX
Erase XX
End Property

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

Private Sub ZZ()
End Sub

Sub BrwRf()
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


