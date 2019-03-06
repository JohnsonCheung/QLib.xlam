Attribute VB_Name = "MDao_Def_Fd_New"
Option Explicit
Public Const EleLblss$ = "*Fld *Ty ?Req ?AlwZLen Dft VTxt VRul TxtSz Expr"
Function FdzStr(FdStr$) As Dao.Field2

End Function
Function Fd(F, Optional Ty As Dao.DataTypeEnum = dbText, Optional Req As Boolean, Optional TxtSz As Byte = 255, Optional ZLen As Boolean, Optional Expr$, Optional Dft$, Optional VRul$, Optional VTxt$) As Dao.Field2
Dim O As New Dao.Field
With O
    .Name = F
    .Required = Req
    If Ty <> 0 Then .Type = Ty
    If Ty = dbText Then
        .Size = TxtSz
        .AllowZeroLength = ZLen
    End If
    If Expr <> "" Then
        CvFd2(O).Expression = Expr
    End If
    O.DefaultValue = Dft
End With
Set Fd = O
End Function

Function FdzBool(F) As Dao.Field2
Set FdzBool = Fd(F, dbBoolean, True, Dft:="0")
End Function

Function FdzByt(F) As Dao.Field2
Set FdzByt = Fd(F, dbByte, True, Dft:="0")
End Function

Function FdzCrtDte(F) As Dao.Field2
Set FdzCrtDte = Fd(F, dbDate, True, Dft:="Now()")
End Function

Function FdzCur(F) As Dao.Field2
Set FdzCur = Fd(F, dbCurrency, True, Dft:="0")
End Function

Function FdzChr(F) As Dao.Field2
Set FdzChr = Fd(F, dbChar, True, Dft:="")
End Function

Function FdzDbl(F) As Dao.Field2
Set FdzDbl = Fd(F, dbDouble, True, Dft:="0")
End Function

Function FdzDte(F) As Dao.Field2
Set FdzDte = Fd(F, dbDate, True, Dft:="0")
End Function

Function FdzDec(F) As Dao.Field2
Set FdzDec = Fd(F, dbDecimal, True, Dft:="0")
End Function

Function FdzEleNmFld(EleNm, F) As Dao.Field2
Dim O As Dao.Field2
Set O = FdzEleNmFld_TNNN(F, EleNm): If Not IsNothing(O) Then Set FdzEleNmFld = O: Exit Function
Select Case EleNm
Case "Nm":  Set FdzEleNmFld = FdzNm(F)
Case "Amt": Set FdzEleNmFld = FdzCur(F): FdzEleNmFld.DefaultValue = 0
Case "Txt": Set FdzEleNmFld = FdzTxt(F, dbText, True): FdzEleNmFld.DefaultValue = """""": FdzEleNmFld.AllowZeroLength = True
Case "Dte": Set FdzEleNmFld = FdzDte(F)
Case "Int": Set FdzEleNmFld = FdzInt(F)
Case "Lng": Set FdzEleNmFld = FdzLng(F)
Case "Dbl": Set FdzEleNmFld = FdzDbl(F)
Case "Sng": Set FdzEleNmFld = FdzSng(F)
Case "Lgc": Set FdzEleNmFld = FdzBool(F)
Case "Mem": Set FdzEleNmFld = FdzMem(F)
End Select
End Function

Private Function FdzEleNmFld_TNNN(F, EleTnnn) As Dao.Field2
If Left(EleTnnn, 1) <> "T" Then Exit Function
Dim A$
A = Mid(EleTnnn, 2)
If CStr(Val(A)) <> A Then Exit Function
Set FdzEleNmFld_TNNN = Fd(F, dbText, True)
With FdzEleNmFld_TNNN
    .Size = A
    .DefaultValue = """"""
    .AllowZeroLength = True
End With
End Function

Function FdzFk(F) As Dao.Field2
Set FdzFk = New Dao.Field
With FdzFk
    .Name = F
    .Type = dbLong
End With
End Function

Function FdzId(F) As Dao.Field2
If Not HasSfx(F, "Id") Then Stop
Dim O As New Dao.Field
With O
    .Name = F
    .Type = dbLong
    .Attributes = Dao.FieldAttributeEnum.dbAutoIncrField
    .Required = True
End With
Set FdzId = O
End Function

Function FdzInt(F) As Dao.Field2
Set FdzInt = Fd(F, dbInteger, True, Dft:="0")
End Function

Function FdzLng(F) As Dao.Field2
Set FdzLng = Fd(F, dbLong, True, Dft:="0")
End Function

Function FdzAtt(F) As Dao.Field2
Set FdzAtt = Fd(F, dbAttachment)
End Function

Function FdzMem(F) As Dao.Field2
Set FdzMem = Fd(F, dbMemo, True, Dft:="""""")
End Function

Function FdzNm(F) As Dao.Field2
If Right(F, 2) <> "Nm" Then Stop
Set FdzNm = Fd(F, dbText, True, 50, False)
End Function

Function FdzPk(F) As Dao.Field2
If Right(F, 2) <> "Id" Then Stop
Set FdzPk = Fd(F, dbLong, True)
FdzPk.Attributes = Dao.FieldAttributeEnum.dbAutoIncrField
End Function

Function FdzSng(F) As Dao.Field2
Set FdzSng = Fd(F, dbSingle, True, Dft:="0")
End Function

Function FdzTim(F) As Dao.Field2
Set FdzTim = Fd(F, dbTime, True, Dft:="0")
End Function

Function FdzShtTys(ShtTys, Fld) As Dao.Field2
'Public Const VdtShtTyLis$ = "ABBytCChrDDteDecILMSTTimTxt"
Dim O As Dao.Field2
Select Case ShtTys
Case "Att", "A":  Set O = FdzAtt(Fld)
Case "Bool", "B": Set O = FdzBool(Fld)
Case "Byt":       Set O = FdzByt(Fld)
Case "Chr", "C":  Set O = FdzCur(Fld)
Case "Dte":       Set O = FdzDte(Fld)
Case "Dec":       Set O = FdzDec(Fld)
Case "Dbl", "D":  Set O = FdzDbl(Fld)
Case "Int", "I":  Set O = FdzInt(Fld)
Case "Lng", "L":  Set O = FdzLng(Fld)
Case "Mem", "M":  Set O = FdzMem(Fld)
Case "Sng", "S":  Set O = FdzSng(Fld)
Case "Txt", "T":  Set O = FdzTxt(Fld)
Case "Tim":       Set O = FdzTim(Fld)
Case Else:
    If FstChr(ShtTys) = "T" Then
        Dim Sz As Byte
        Sz = CByte(RmvFstChr(ShtTys))
        Set O = FdzTxt(Fld, Sz)
        Exit Function
    End If
    ThwShtTyEr CSub, ShtTys
End Select
Set FdzShtTys = O
End Function

Function FdzStdFld(StdFld, Optional T$) As Dao.Field2
Dim R2$, R3$: R2 = Right(StdFld, 2): R3 = Right(StdFld, 3)
Select Case True
Case StdFld = "CrtDte": Set FdzStdFld = FdzCrtDte(StdFld)
Case T & "Id" = StdFld: Set FdzStdFld = FdzPk(StdFld)
Case R2 = "Id":    Set FdzStdFld = FdzFk(StdFld)
Case R2 = "Ty":    Set FdzStdFld = FdzTy(StdFld)
Case R2 = "Nm":    Set FdzStdFld = FdzNm(StdFld)
Case R3 = "Dte":   Set FdzStdFld = FdzDte(StdFld)
Case R3 = "Amt":   Set FdzStdFld = FdzCur(StdFld)
Case R3 = "Att":   Set FdzStdFld = FdzAtt(StdFld)
End Select
End Function

Private Sub Z_FdzFdStr()
Dim Act As Dao.Field2, Ept As Dao.Field2, mFdStr
mFdStr = "AA Int Req AlwZLen Dft=ABC TxtSz=10"
Set Ept = New Dao.Field
With Ept
    .Type = Dao.DataTypeEnum.dbInteger
    .Name = "AA"
    '.AllowZeroLength = False
    .DefaultValue = "ABC"
    .Required = True
    .Size = 2
End With
GoSub Tst
Exit Sub
Tst:
    Set Act = FdzFdStr(mFdStr)
    If Not IsEqFd(Act, Ept) Then
        D LyzMsgNap("Act", "FdStr", FdStr(Act))
        D LyzMsgNap("Ept", "FdStr", FdStr(Ept))
    End If
    Return
End Sub

Private Sub Z_FdzFdStr1()
Dim FdStr$
FdStr = "Txt Req Dft=ABC [VTxt=Loc must cannot be blank] [VRul=IsNull([Loc]) or Trim(Loc)='']"
GoSub Tst
Exit Sub
Tst:
    Set Act = FdzFdStr(FdStr)
    Stop
    Return
End Sub

Function FdzFdStr(FdStr) As Dao.Field2
Dim Fld$, TyStr$, Req As Boolean, AlwZLen As Boolean, Dft$, VTxt$, VRul$, TxtSz As Byte, Expr$
Dim L$: L = FdStr
Dim Vy(): Vy = VyzLinLbl(L, EleLblss)
AsgAp Vy, _
    Fld, TyStr, Req, AlwZLen, Dft, VTxt, VRul, TxtSz, Expr
Set FdzFdStr = Fd( _
    Fld, DaoTyzShtTy(TyStr), Req, TxtSz, AlwZLen, Expr, Dft, VRul, VTxt)
End Function

Function FdzTxt(F, Optional TxtSz As Byte = 255, Optional ZLen As Boolean, Optional Expr$, Optional Dft$, Optional Req As Boolean, Optional VRul$, Optional VTxt$) As Dao.Field2
Set FdzTxt = Fd(F, dbText, Req, TxtSz, ZLen, Expr, Dft, VRul, VTxt)
End Function

Function FdzTy(F) As Dao.Field2
Set FdzTy = Fd(F, dbText, True, 20, ZLen:=False)
End Function

Private Sub ZZ()
Dim A As Variant
Dim B As Dao.DataTypeEnum
Dim C As Boolean
Dim D As Byte
Dim E$
FdzCrtDte A
FdzCur A
FdzDte A
FdzEleNmFld A, A
FdzFk A
FdzId A
FdzNm A
FdzPk A
FdzFdStr A
FdzTxt A, D, C, E, E, C, E, E
FdzTy A
End Sub

Private Sub Z()
End Sub
