Attribute VB_Name = "MIde_Ens_SubZZ"
Option Explicit
Const CMod$ = "MIde_Ens_SubZZ."

Function SubZZEpt$(A As CodeModule) ' SubZZ is Sub ZZ() bodyLines
'Sub ZZ() has all calling of public method with dummy parameter so that it can Shf-F2
Dim mPubMthLinAy$()        ' Mth mPubMthLinAy PUB
Dim mPubPrpGetNm As Aset
Dim mPubMthPmAy$()         ' From mPubMthLinAy    mPubMthLinAy & mPm same sz     ' mPubMthPmAy is the string in the bracket of the MthmPubMthLinAyLin
Dim mMthNyPub$()      ' From mPubMthLinAy    mPubMthLinAy & mMthNm same sz
Dim mArgAy$()         ' Each mArgAy in mArgAy become on mPubMthPmAy   Eg, 1-mArgAy = A$, B$, C%, D As XYZ => 4-mPubMthPmAy
                     ' ArgSfxDic is Key=ArgSfx and Val=A, B, C
                     ' ArgSfx is mArgAy-without-Nm
Dim ArgSfx$()
Dim ArgAset As Aset
Dim ArgSfxToAbcDic As Dictionary
Dim mCallingPmAy$()
Dim mHasPrp As Boolean

    mPubMthLinAy = MthLinAyzPub(Src(A))
    Set mPubPrpGetNm = WPrpGetAset(mPubMthLinAy)
    mPubMthPmAy = AyTakBetBkt(mPubMthLinAy) ' Each Mth return an-Ele to call
    mMthNyPub = MthNyzMthLinAy(mPubMthLinAy)
    mArgAy = ArgAyzPmAy(mPubMthPmAy)
    ArgSfx = ArgSfxAy(mArgAy)
    Set ArgAset = AsetzAy(ArgSfx)
    Set ArgSfxToAbcDic = ArgAset.AbcDic
    mCallingPmAy = WCallingPmAy(mPubMthPmAy, ArgSfxToAbcDic)
    mHasPrp = mPubPrpGetNm.Cnt > 0
'-------------
Dim mDimLy$()
Dim mCallingLy$()
    mDimLy = WDimLy(ArgSfxToAbcDic, mHasPrp)   ' 1-mArgAy => 1-DimLin
    mCallingLy = WCallingLy(mMthNyPub, mPubMthPmAy, ArgSfxToAbcDic, mPubPrpGetNm)
D mDimLy
Stop
Dim O$()
    PushI O, "Private Sub ZZ()"
    PushIAy O, mDimLy
    PushIAy O, mCallingLy
    PushI O, "End Sub"
SubZZEpt = JnCrLf(O)
End Function

Function ArgAyzPmAy(PmAy$()) As String()
Dim Pm, Arg
For Each Pm In Itr(PmAy)
    For Each Arg In Itr(SplitCommaSpc(Pm))
        PushI ArgAyzPmAy, Arg
    Next
Next
End Function

Private Function ArgSfxAy(ArgAy$()) As String()
Dim Arg
For Each Arg In Itr(ArgAy)
    PushI ArgSfxAy, ArgSfx(Arg)
Next
End Function

Private Function WCallingLin$(MthNm, CallingPm$, PrpGetAset As Aset)
If PrpGetAset.Has(MthNm) Then
    WCallingLin = "XX = " & MthNm & "(" & CallingPm & ")"  ' The MthNm is object, no need to add [Set] XX =, the compiler will not check for this
Else
    WCallingLin = MthNm & AddPfxSpc_IfNonBlank(CallingPm)
End If
End Function

Private Function WCallingLy(MthNy$(), PmAy$(), ArgDic As Dictionary, PrpGetAset As Aset) As String()
'A$() & PmAy$() are same sz
'ArgDic: Key is ArgSfx(Arg-without-Name), Val is A,B,..
'CallingLin is {MthNm} A,B,C,...
'PrpGetAset    is PrpNm set
Dim MthNm, CallingPm$, Pm$, J%, O$()
For Each MthNm In Itr(MthNy)
    Pm = PmAy(J)
    CallingPm = WCallingPm(Pm, ArgDic)
    PushI O, WCallingLin(MthNm, CallingPm, PrpGetAset)
    J = J + 1
Next
WCallingLy = AyQSrt(O)
End Function

Private Function WCallingPm$(Pm, ArgDic As Dictionary)
Dim O$(), Arg
For Each Arg In Itr(AyTrim(SplitComma(Pm)))
    PushI O, ArgDic(ArgSfx(Arg))
Next
'WCallingPm = CommaSpc(O)
End Function

Private Function WCallingPmAy(PmAy$(), ArgDic As Dictionary) As String()
Dim Pm
For Each Pm In Itr(PmAy)
    PushI WCallingPmAy, WCallingPm(Pm, ArgDic)
Next
End Function

Private Function WDimLy(ArgDic As Dictionary, HasPrp As Boolean) As String()  '1-Arg => 1-DimLin
Dim ArgSfx, S$
For Each ArgSfx In ArgDic.Keys
    If HasPfx(ArgSfx, "As ") Then
        S = " "
    Else
        S = ""
    End If
    PushI WDimLy, "Dim " & ArgDic(ArgSfx) & S & ArgSfx
Next
If HasPrp Then PushI WDimLy, "Dim XX"
End Function

Private Function WPrpGetAset(MthDclAy$()) As Aset
Dim Lin, O As Aset
Set O = EmpAset
For Each Lin In Itr(MthDclAy)
'    If IsPrpLin(Lin) Then AsetPush O, MthNm(Lin)
Next
Set WPrpGetAset = O
End Function

Private Sub Z_SubZZEpt()
Dim M As CodeModule
GoSub T1
'GoSub T2
Exit Sub
T1:
    Set M = Md("MDamZ_Db_Dbt")
    GoTo Tst
T2:
    Set M = CurMd
    GoTo Tst
Tst:
    Act = SubZZEpt(M)
    Brw Act
    Stop
    C
Return
End Sub

Private Sub ZZ()
Dim A
Dim B As CodeModule
Dim C$()
Dim XX
ArgSfx A
SubZZEpt B
MthNyzMthLinAy C
End Sub

Private Sub Z()
'MIde_Ens_SubZZ.ArgSfx
Z_SubZZEpt
End Sub

Private Property Get Z_SubZZzMd__Ept2$()
Const A_1$ = "Private Sub ZZ()" & _
vbCrLf & "Dim A As Variant" & _
vbCrLf & "Dim B As Dictionary" & _
vbCrLf & "Dim C() As Dictionary" & _
vbCrLf & "Dim D$" & _
vbCrLf & "Dim E$()" & _
vbCrLf & "Dim F As Boolean" & _
vbCrLf & "Dim G()" & _
vbCrLf & "CvDic A" & _
vbCrLf & "CvDicAy A" & _
vbCrLf & "DicAyAdd B, C" & _
vbCrLf & "AddDicKeyPfx B, A" & _
vbCrLf & "DicAddOrUpd B, D, A, D" & _
vbCrLf & "DicAllKeyIsNm B" & _
vbCrLf & "DicAllKeyIsStr B" & _
vbCrLf & "DicAllValIsStr B" & _
vbCrLf & "DicAyKy C" & _
vbCrLf & "DicByDry A" & _
vbCrLf & "DicClone B" & _
vbCrLf & "DrDicKy B, E"

Const A_2$ = "DicFny F" & _
vbCrLf & "DicIntersect B, B" & _
vbCrLf & "IsDiczEmp B" & _
vbCrLf & "IsEqDic B, B" & _
vbCrLf & "ThwDifDic B, B, D, D, D" & _
vbCrLf & "IsDiczLines B" & _
vbCrLf & "IsDiczStr B" & _
vbCrLf & "KeySyzDic B" & _
vbCrLf & "FmtDic1 B" & _
vbCrLf & "ValOfDicKyJn B, A, D" & _
vbCrLf & "SyzDicKy B, E" & _
vbCrLf & "FmtDicTit B, D" & _
vbCrLf & "LineszDic B" & _
vbCrLf & "FmtDic11 B" & _
vbCrLf & "FmtDic2 B" & _
vbCrLf & "FmtDic2__1 D, D" & _
vbCrLf & "DicMap B, D" & _
vbCrLf & "DicMaxValSz B" & _
vbCrLf & "DicMge B, D, G" & _
vbCrLf & "DicMinus B, B"

Const A_3$ = "DicSelIntozAy B, E" & _
vbCrLf & "DicSelIntoSy B, E" & _
vbCrLf & "SyzDicKey B" & _
vbCrLf & "DiczSwapKV B" & _
vbCrLf & "DicTy B" & _
vbCrLf & "DicTyBrw B" & _
vbCrLf & "DicValOpt B, A" & _
vbCrLf & "KeyzLikssDic_Itm B, A" & _
vbCrLf & "MapStrDic A" & _
vbCrLf & "MayDicValOpt B, A" & _
vbCrLf & "End Sub"

Z_SubZZzMd__Ept2 = A_1 & vbCrLf & A_2 & vbCrLf & A_3
End Property

Private Sub Z_SubZZEptzMd()
Dim A As CodeModule
GoSub Cas2
GoSub Cas1
Exit Sub
Cas2:
    Set A = Md("MVb_Dic")
    UpdConst "SubZZEptzMd_Ept2", SubZZEpt(A): Return
    Ept = Z_SubZZzMd__Ept2
    GoSub Tst
    Return
Cas1:
    Set A = CurMd
    'UpdMdConstValOfFt "Z_SubZZEptMd_1", SubZZEpt(A): Return
    'Ept = Z_SubZZEptMd_1
    GoSub Tst
    Return
Tst:
    'Act = SubZZEptzMd(A)
    C
    Return
End Sub


