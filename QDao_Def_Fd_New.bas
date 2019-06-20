Attribute VB_Name = "QDao_Def_Fd_New"
Option Compare Text
Option Explicit
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_Def_Fd_New."
Public Const EleLblss$ = "*Fld *Ty ?Req ?AlwZLen Dft VTxt VRul TxtSz Expr"

Function FdzStr(FdStr$) As DAO.Field2
End Function

Function Fd(F$, Optional Ty As DAO.DataTypeEnum = dbText, Optional Req As Boolean, Optional TxtSz As Byte = 255, Optional ZLen As Boolean, Optional Expr$, Optional Dft$, Optional VRul$, Optional VTxt$) As DAO.Field2
Dim O As New DAO.Field
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

Function FdzBool(F$) As DAO.Field2
Set FdzBool = Fd(F, dbBoolean, True, Dft:="0")
End Function

Function FdzByt(F$) As DAO.Field2
Set FdzByt = Fd(F, dbByte, True, Dft:="0")
End Function

Function FdzCrtDte(F$) As DAO.Field2
Set FdzCrtDte = Fd(F, dbDate, True, Dft:="Now()")
End Function

Function FdzCur(F$) As DAO.Field2
Set FdzCur = Fd(F, dbCurrency, True, Dft:="0")
End Function

Function FdzChr(F$) As DAO.Field2
Set FdzChr = Fd(F, dbChar, True, Dft:="")
End Function

Function FdzDbl(F$) As DAO.Field2
Set FdzDbl = Fd(F, dbDouble, True, Dft:="0")
End Function

Function FdzDte(F$) As DAO.Field2
Set FdzDte = Fd(F, dbDate, True, Dft:="0")
End Function

Function FdzDec(F$) As DAO.Field2
Set FdzDec = Fd(F, dbDecimal, True, Dft:="0")
End Function

Function FdzEle(Ele$, F$) As DAO.Field2
Dim O As DAO.Field2
Set O = FdzTnnn(F, Ele): If Not IsNothing(O) Then Set FdzEle = O: Exit Function
Select Case Ele
Case "Nm":  Set FdzEle = FdzNm(F)
Case "Amt": Set FdzEle = FdzCur(F): FdzEle.DefaultValue = 0
Case "Txt": Set FdzEle = FdzTxt(F, dbText, True): FdzEle.DefaultValue = """""": FdzEle.AllowZeroLength = True
Case "Dte": Set FdzEle = FdzDte(F)
Case "Int": Set FdzEle = FdzInt(F)
Case "Lng": Set FdzEle = FdzLng(F)
Case "Dbl": Set FdzEle = FdzDbl(F)
Case "Sng": Set FdzEle = FdzSng(F)
Case "Lgc": Set FdzEle = FdzBool(F)
Case "Mem": Set FdzEle = FdzMem(F)
End Select
End Function

Private Function FdzTnnn(F$, EleTnnn) As DAO.Field2
If Left(EleTnnn, 1) <> "T" Then Exit Function
Dim A$
A = Mid(EleTnnn, 2)
If CStr(Val(A)) <> A Then Exit Function
Set FdzTnnn = Fd(F, dbText, True)
With FdzTnnn
    .Size = A
    .DefaultValue = """"""
    .AllowZeroLength = True
End With
End Function

Function FdzFk(F$) As DAO.Field2
Set FdzFk = New DAO.Field
With FdzFk
    .Name = F
    .Type = dbLong
End With
End Function

Function FdzId(F$) As DAO.Field2
If Not HasSfx(F$, "Id") Then Thw CSub, "FldNm must has Sfx-Id", "FldNm", F
Dim O As New DAO.Field
With O
    .Name = F
    .Type = dbLong
    .Attributes = DAO.FieldAttributeEnum.dbAutoIncrField
    .Required = True
End With
Set FdzId = O
End Function

Function FdzInt(F$) As DAO.Field2
Set FdzInt = Fd(F, dbInteger, True, Dft:="0")
End Function

Function FdzLng(F$) As DAO.Field2
Set FdzLng = Fd(F, dbLong, True, Dft:="0")
End Function

Function FdzAtt(F$) As DAO.Field2
Set FdzAtt = Fd(F, dbAttachment)
End Function

Function FdzMem(F$) As DAO.Field2
Set FdzMem = Fd(F, dbMemo, True, Dft:="""""")
End Function

Function FdzNm(F$) As DAO.Field2
If Right(F, 2) <> "Nm" Then Stop
Set FdzNm = Fd(F, dbText, True, 50, False)
End Function

Function FdzPk(F$) As DAO.Field2
If Right(F, 2) <> "Id" Then Stop
Set FdzPk = Fd(F, dbLong, True)
FdzPk.Attributes = DAO.FieldAttributeEnum.dbAutoIncrField
End Function

Function FdzSng(F$) As DAO.Field2
Set FdzSng = Fd(F, dbSingle, True, Dft:="0")
End Function

Function FdzTim(F$) As DAO.Field2
Set FdzTim = Fd(F, dbTime, True, Dft:="0")
End Function

Function FdzShtTys(ShtTys$, F$) As DAO.Field2
Const CSub$ = CMod & "FdzShtTys"
'Public Const ShtTyLis$ = "ABBytCChrDDteDecILMSTTimTxt"
Dim O As DAO.Field2
Select Case ShtTys
Case "Att", "A":  Set O = FdzAtt(F)
Case "Bool", "B": Set O = FdzBool(F)
Case "Byt":       Set O = FdzByt(F)
Case "Chr", "C":  Set O = FdzCur(F)
Case "Dte":       Set O = FdzDte(F)
Case "Dec":       Set O = FdzDec(F)
Case "Dbl", "D":  Set O = FdzDbl(F)
Case "Int", "I":  Set O = FdzInt(F)
Case "Lng", "L":  Set O = FdzLng(F)
Case "Mem", "M":  Set O = FdzMem(F)
Case "Sng", "S":  Set O = FdzSng(F)
Case "Txt", "T":  Set O = FdzTxt(F)
Case "Tim":       Set O = FdzTim(F)
Case Else:
    If FstChr(ShtTys) = "T" Then
        Dim Si As Byte
        Si = CByte(RmvFstChr(ShtTys))
        Set O = FdzTxt(F, Si)
        Exit Function
    End If
    Thw CSub, "ShtTys Err", "ShtTys", ShtTys
End Select
Set FdzShtTys = O
End Function

Function FdzStdFldNm(StdFldNm$, Optional T) As DAO.Field2
Dim R2$, R3$: R2 = Right(StdFldNm$, 2): R3 = Right(StdFldNm, 3)
Dim O As DAO.Field2
Select Case True
Case StdFldNm = "CrtDte": Set O = FdzCrtDte(StdFldNm)
Case T & "Id" = StdFldNm: Set O = FdzPk(StdFldNm)
Case R2 = "Id":    Set O = FdzFk(StdFldNm)
Case R2 = "Ty":    Set O = Fd(StdFldNm, dbText, True, 20, ZLen:=False)
Case R2 = "Nm":    Set O = FdzNm(StdFldNm)
Case R3 = "Dte":   Set O = FdzDte(StdFldNm)
Case R3 = "Amt":   Set O = FdzCur(StdFldNm)
Case R3 = "Att":   Set O = FdzAtt(StdFldNm)
End Select
Set FdzStdFldNm = O
End Function

Private Sub Z_FdzFdStr()
Dim Act As DAO.Field2, Ept As DAO.Field2, mFdStr$
mFdStr = "AA Int Req AlwZLen Dft=ABC TxtSz=10"
Set Ept = New DAO.Field
With Ept
    .Type = DAO.DataTypeEnum.dbInteger
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

Function FdzFdStr(FdStr$) As DAO.Field2
Dim Fld$, TyStr$, Req As Boolean, AlwZLen As Boolean, Dft$, VTxt$, VRul$, TxtSz As Byte, Expr$
Dim L$: L = FdStr
Dim Vy(): Vy = ShfVy(L, EleLblss)
AsgAp Vy, _
    Fld, TyStr, Req, AlwZLen, Dft, VTxt, VRul, TxtSz, Expr
Set FdzFdStr = Fd( _
    Fld, DaoTyzShtTy(TyStr), Req, TxtSz, AlwZLen, Expr, Dft, VRul, VTxt)
End Function

Function FdzTxt(F$, Optional TxtSz As Byte = 255, Optional ZLen As Boolean, Optional Expr$, Optional Dft$, Optional Req As Boolean, Optional VRul$, Optional VTxt$) As DAO.Field2
Set FdzTxt = Fd(F, dbText, Req, TxtSz, ZLen, Expr, Dft, VRul, VTxt)
End Function

Private Sub Z()
Dim A As Variant
Dim B As DAO.DataTypeEnum
Dim C As Boolean
Dim D As Byte
Dim E$
End Sub



