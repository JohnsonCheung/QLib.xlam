Attribute VB_Name = "QIde_Ens_SubZZ"
Option Compare Text
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Ens_SubZZ."

Function SubZZEpt$(A As CodeModule) ' SubZZ is Sub ZZ() bodyLines
'Sub ZZ() has all calling of public method with dummy parameter so that it can Shf-F2
Dim mPMthLiny$()        ' Mth mPMthLiny PUB
Dim mPubPrpGetNm As Aset
Dim mPMthPmAy$()         ' From mPMthLiny    mPMthLiny & mPm same sz     ' mPMthPmAy is the string in the bracket of the MthmPMthLinyLin
Dim mMthnyPub$()      ' From mPMthLiny    mPMthLiny & mMthn same sz
Dim mArgAy$()         ' Each mArgAy in mArgAy become on mPMthPmAy   Eg, 1-mArgAy = A$, B$, C%, D As XYZ => 4-mPMthPmAy
                     ' ArgSfxDic is Key=ArgSfx and Val=A, B, C
                     ' ArgSfx is mArgAy-without-Nm
Dim ArgSfx$()
Dim ArgAset As Aset
Dim ArgSfxToAbcDic As Dictionary
Dim mCallingPmAy$()
Dim mHasPrp As Boolean

    mPMthLiny = MthLinyzPub(Src(A))
    Set mPubPrpGetNm = WPrpGetAset(mPMthLiny)
    mPMthPmAy = SyTakBetBkt(mPMthLiny) ' Each Mth return an-Ele to call
    mMthnyPub = MthnyzMthLiny(mPMthLiny)
    mArgAy = ArgAyzPmAy(mPMthPmAy)
    ArgSfx = ArgSfxy(mArgAy)
    Set ArgAset = AsetzAy(ArgSfx)
    Set ArgSfxToAbcDic = ArgAset.AbcDic
    mCallingPmAy = WCallingPmAy(mPMthPmAy, ArgSfxToAbcDic)
    mHasPrp = mPubPrpGetNm.Cnt > 0
'-------------
Dim mDimLy$()
Dim mCallingLy$()
    mDimLy = WDimLy(ArgSfxToAbcDic, mHasPrp)   ' 1-mArgAy => 1-DimLin
    mCallingLy = WCallingLy(mMthnyPub, mPMthPmAy, ArgSfxToAbcDic, mPubPrpGetNm)
D mDimLy
Stop
Dim O$()
    PushI O, "Private Sub ZZ()"
    PushIAy O, mDimLy
    PushIAy O, mCallingLy
    PushI O, "End Sub"
SubZZEpt = JnCrLf(O)
End Function

Private Function WCallingLin(Mthn, CallingPm$, PrpGetAset As Aset)
If PrpGetAset.Has(Mthn) Then
    WCallingLin = "XX = " & Mthn & "(" & CallingPm & ")"  ' The Mthn is object, no need to add [Set] XX =, the compiler will not check for this
Else
    WCallingLin = Mthn & AddPfxSpczIfNonBlank(CallingPm)
End If
End Function

Private Function WCallingLy(Mthny$(), PmAy$(), ArgDic As Dictionary, PrpGetAset As Aset) As String()
'A$() & PmAy$() are same sz
'ArgDic: Key is ArgSfx(Arg-without-Name), Val is A,B,..
'CallingLin is {Mthn} A,B,C,...
'PrpGetAset    is PrpNm set
Dim Mthn, CallingPm$, Pm$, J%, O$()
For Each Mthn In Itr(Mthny)
    Pm = PmAy(J)
    CallingPm = WCallingPm(Pm, ArgDic)
    PushI O, WCallingLin(Mthn, CallingPm, PrpGetAset)
    J = J + 1
Next
WCallingLy = QSrt1(O)
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
'    If IsPrpLin(Lin) Then AsetPush O, Mthn(Lin)
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
    Set M = CMd
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
MthnyzMthLiny C
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
vbCrLf & "AddDicAy B, C" & _
vbCrLf & "AddDicKeyPfx B, A" & _
vbCrLf & "DicAddOrUpd B, D, A, D" & _
vbCrLf & "DicAllKeyIsNm B" & _
vbCrLf & "DicAllKeyIsStr B" & _
vbCrLf & "DicAllValIsStr B" & _
vbCrLf & "DicAyKy C" & _
vbCrLf & "DicByDry A" & _
vbCrLf & "CloneDic B" & _
vbCrLf & "DrDicKy B, E"

Const A_2$ = "DicFny F" & _
vbCrLf & "DicIntersectAy B, B" & _
vbCrLf & "IsEmpDic B" & _
vbCrLf & "IsEqDic B, B" & _
vbCrLf & "ThwIf_DifDic B, B, D, D, D" & _
vbCrLf & "IsDicOfLines B" & _
vbCrLf & "IsDicOfStr B" & _
vbCrLf & "KeySyzDic B" & _
vbCrLf & "FmtDic1 B" & _
vbCrLf & "ValzDicIfKyJn B, A, D" & _
vbCrLf & "SyzDicKy B, E" & _
vbCrLf & "FmtDicTit B, D" & _
vbCrLf & "LineszDic B" & _
vbCrLf & "FmtDic11 B" & _
vbCrLf & "FmtDic2 B" & _
vbCrLf & "FmtDic2__1 D, D" & _
vbCrLf & "DicMap B, D" & _
vbCrLf & "MaxSizAyDic B" & _
vbCrLf & "MgeDic B, D, G" & _
vbCrLf & "MinusDic B, B"

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
    'UpdConst "SubZZEptzMd_Ept2", SubZZEpt(A): Return
    Ept = Z_SubZZzMd__Ept2
    GoSub Tst
    Return
Cas1:
    Set A = CMd
    'UpdMdCnstBrkzFt "Z_SubZZEptMd_1", SubZZEpt(A): Return
    'Ept = Z_SubZZEptMd_1
    GoSub Tst
    Return
Tst:
    'Act = SubZZEptzMd(A)
    C
    Return
End Sub


