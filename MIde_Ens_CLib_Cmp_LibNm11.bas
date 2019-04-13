Attribute VB_Name = "MIde_Ens_CLib_Cmp_LibNm11"
Option Explicit
Enum eLibNmTy
    eeByDic
    eeByFstCml
    eeByPjNm
End Enum

Function NoLibMdNy() As String()
Dim I
For Each I In CmpAy
    If LibNm(CvCmp(I), eeByDic) = "" Then
        PushI NoLibMdNy, CvCmp(I).Name
    End If
Next
End Function

Function LibNm$(A As VBComponent, B As eLibNmTy)
LibNm = StrBef(A.Name, "_")
End Function

Function MdNmToLibNmDic() As Dictionary
Static D As Dictionary: If IsNothing(D) Then Set D = MdNmToLibNmDiczPj(CurPj)
Set MdNmToLibNmDic = D
End Function

Private Function MdNmToLibNmDiczPj(A As VBProject) As Dictionary
Dim I, D1 As Dictionary, D2 As Dictionary
Set D1 = MdNmToLibNmDiczDef
Set D2 = MdPfxToLibNmDic
Dim O As New Dictionary
For Each I In MdNyzPj(A)
    Select Case True
    Case D1.Exists(I):          O.Add I, D1(I)
    Case D2.Exists(MdPfx(I)):   O.Add I, D2(MdPfx(I))
    Case Else: O.Add I, ""
    End Select
Next
Set MdNmToLibNmDiczPj = O
End Function

Private Function MdPfxToLibNmDic() As Dictionary
Static A As Dictionary
If IsNothing(A) Then Set A = Dic(C_MdPfxToLibNmLy)
Set MdPfxToLibNmDic = A
End Function

Private Function MdNmToLibNmDiczDef() As Dictionary
Static A As Dictionary
If IsNothing(A) Then Set A = Dic(C_MdNmToLibNmLy)
Set MdNmToLibNmDiczDef = A
End Function

Function C_MdNmToLibNmLy() As String()
Erase xx
X "MCmpAdd QIde"
X "MUSysRegInf QIde"
X "MLinShould QIde"
X "MMdyPj QIde"
X "MTreeWs QXls"
X "SyPair QVb"
X "MGoWsLnk QXls"
X "MSyDic QVb"
X "MClrCell QXls"
X "MTreeWsInstall QXls"
X "MDefDic QVb"
X "MSepLin QDta"
C_MdNmToLibNmLy = xx
Erase xx
End Function

Function C_MdPfxToLibNmLy() As String()
Erase xx
X "MCml QVb"
X "Act QIde"
X "MCmp QIde"
X "MWs QXls"
X "MWs QXls"
X "Arg QIde"
X "Aset QVb"
X "Pj QIde"
X "MPub QIde"
X "MItp QVb"
X "AShp QShpCst"
X "Blk QTp"
X "Drs QDta"
X "Ds QDta"
X "Dt QDta"
X "FTIx QVb"
X "Gp QTp"
X "Li QDao"
X "Lid QDao"
X "Lin QVb"
X "Lnx QVb"
X "Lt QDao"
X "MAcs QAcs"
X "MAdo QDao"
X "MApp QApp"
X "MAy QVb"
X "MBtn QVb"
X "MChk QDao"
X "MCrypto QVb"
X "Md QIde"
X "MDao QDao"
X "MDb QDao"
X "MDbq QDao"
X "MDbt QDao"
X "MDic QVb"
X "MDta QDta"
X "MEns QIde"
X "MFs QVb"
X "MIde QIde"
X "MLi QDao"
X "MLid QDao"
X "MLo QXls"
X "MMd QIde"
X "MMth QIde"
X "Module1 QVb"
X "MPj QIde"
X "MRpt QShpCst"
X "MSql QDao"
X "MSsk QDao"
X "MStr QVb"
X "Mth QIde"
X "MTp QTp"
X "MTst QVb"
X "MVb QVb"
X "MAp QVb"
X "MLib QIde"
X "MXls QXls"
X "Rel QVb"
X "RRCC QVb"
X "S1 QVb"
X "Sw QTp"
X "Vbe QIde"
X "Wh QVb"
X "MPth QVb"
X "MInto QVb"
X "Lof QXls"
X "MFill QXls"
X "MExpand QVb"
X "MWb QXls"
X "MRun QVb"
X "MWc QXls"
X "MArg QIde"
X "MNMth QIde"
X "MNCmp QIde"
X "Cmp QIde"
C_MdPfxToLibNmLy = xx
Erase xx
End Function

Function LibDef() As String()
LibDef = _
FmtAyT2(AyAdd( _
AyAddPfx(C_MdNmToLibNmLy, "MdNm "), _
AyAddPfx(C_MdPfxToLibNmLy, "MdPfx ")))
End Function
