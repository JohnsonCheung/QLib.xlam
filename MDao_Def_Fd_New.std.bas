Attribute VB_Name = "MDao_Def_Fd_New"
Option Explicit

Function NewFd(F, Optional Ty As DAO.DataTypeEnum = dbText, Optional Req As Boolean, Optional TxtSz As Byte = 255, Optional ZLen As Boolean, Optional Expr$, Optional Dft$, Optional VRul$, Optional VTxt$) As DAO.Field2
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
Set NewFd = O
End Function

Function FdzBool(F) As DAO.Field2
Set FdzBool = NewFd(F, dbBoolean, True, Dft:="0")
End Function

Function FdzCrtDte(F) As DAO.Field2
Set FdzCrtDte = NewFd(F, dbDate, True, Dft:="Now()")
End Function

Function FdzCur(F) As DAO.Field2
Set FdzCur = NewFd(F, dbCurrency, True, Dft:="0")
End Function

Function FdzDbl(F) As DAO.Field2
Set FdzDbl = NewFd(F, dbDouble, True, Dft:="0")
End Function

Function FdzDte(F) As DAO.Field2
Set FdzDte = NewFd(F, dbDate, True, Dft:="0")
End Function

Function FdzEleNmFld(EleNm, F) As DAO.Field2
Dim O As DAO.Field2
Set O = FdzEleNmFld_TNNN(F, EleNm): If Not IsNothing(O) Then Set FdzEleNmFld = O: Exit Function
Select Case EleNm
Case "Nm":  Set FdzEleNmFld = FdzNm(F)
Case "Amt": Set FdzEleNmFld = FdzCur(F): FdzEleNmFld.DefaultValue = 0
Case "Txt": Set FdzEleNmFld = FdzTxt(F, dbText, True): FdzEleNmFld.DefaultValue = """""": FdzEleNmFld.AllowZeroLength = True
Case "Dte": Set FdzEleNmFld = FdzDte(F)
Case "Int": Set FdzEleNmFld = NewFdINT(F)
Case "Lng": Set FdzEleNmFld = NewFdzLng(F)
Case "Dbl": Set FdzEleNmFld = FdzDbl(F)
Case "Sng": Set FdzEleNmFld = FdzSng(F)
Case "Lgc": Set FdzEleNmFld = FdzBool(F)
Case "Mem": Set FdzEleNmFld = FdzMem(F)
End Select
End Function

Private Function FdzEleNmFld_TNNN(F, EleTnnn) As DAO.Field2
If Left(EleTnnn, 1) <> "T" Then Exit Function
Dim A$
A = Mid(EleTnnn, 2)
If CStr(Val(A)) <> A Then Exit Function
Set FdzEleNmFld_TNNN = NewFd(F, dbText, True)
With FdzEleNmFld_TNNN
    .Size = A
    .DefaultValue = """"""
    .AllowZeroLength = True
End With
End Function

Function FdzFk(F) As DAO.Field2
Set FdzFk = New DAO.Field
With FdzFk
    .Name = F
    .Type = dbLong
End With
End Function

Function FdzId(F) As DAO.Field2
If Not HasSfx(F, "Id") Then Stop
Dim O As New DAO.Field
With O
    .Name = F
    .Type = dbLong
    .Attributes = DAO.FieldAttributeEnum.dbAutoIncrField
    .Required = True
End With
Set FdzId = O
End Function

Function NewFdINT(F) As DAO.Field2
Set NewFdINT = NewFd(F, dbInteger, True, Dft:="0")
End Function

Function NewFdzLng(F) As DAO.Field2
Set NewFdzLng = NewFd(F, dbLong, True, Dft:="0")
End Function

Function FdzAtt(F) As DAO.Field2
Set FdzAtt = NewFd(F, dbAttachment)
End Function

Function FdzMem(F) As DAO.Field2
Set FdzMem = NewFd(F, dbMemo, True, Dft:="""""")
End Function

Function FdzNm(F) As DAO.Field2
If Right(F, 2) <> "Nm" Then Stop
Set FdzNm = NewFd(F, dbText, True, 50, False)
End Function

Function FdzPk(F) As DAO.Field2
If Right(F, 2) <> "Id" Then Stop
Set FdzPk = NewFd(F, dbLong, True)
FdzPk.Attributes = DAO.FieldAttributeEnum.dbAutoIncrField
End Function

Function FdzSng(F) As DAO.Field2
Set FdzSng = NewFd(F, dbSingle, True, Dft:="0")
End Function

Function StdFd(F, Optional T$) As DAO.Field2
Dim R2$, R3$: R2 = Right(F, 2): R3 = Right(F, 3)
Select Case True
Case F = "CrtDte": Set StdFd = FdzCrtDte(F)
Case T & "Id" = F: Set StdFd = FdzPk(F)
Case R2 = "Id":    Set StdFd = FdzId(F)
Case R2 = "Ty":    Set StdFd = NewFdTY(F)
Case R2 = "Nm":    Set StdFd = FdzNm(F)
Case R3 = "Dte":   Set StdFd = FdzDte(F)
Case R3 = "Amt":   Set StdFd = FdzCur(F)
Case R3 = "Att":   Set StdFd = FdzAtt(F)
End Select
End Function

Private Sub Z_FdzFdStr()
Dim Act As DAO.Field2, Ept As DAO.Field2, mFdStr
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
    Return
End Sub

Function FdzFdStr(FdStr) As DAO.Field2
Const EleLblss$ = "*Fld *Ty ?Req ?AlwZLen Dft VTxt VRul TxtSz Expr"
Dim Fld$, TyStr$, Req As Boolean, AlwZLen As Boolean, Dft$, VTxt$, VRul$, TxtSz As Byte, Expr$
Dim L$: L = FdStr
Dim VyzDicKK(): VyzDicKK = ShfVal(L, EleLblss)
AsgApAy VyzDicKK, _
    Fld, TyStr, Req, AlwZLen, Dft, VTxt, VRul, TxtSz, Expr
Set FdzFdStr = NewFd( _
    Fld, DaoTyzShtTy(TyStr), Req, TxtSz, AlwZLen, Expr, Dft, VRul, VTxt)
End Function

Function FdzTxt(F, Optional TxtSz As Byte = 255, Optional ZLen As Boolean, Optional Expr$, Optional Dft$, Optional Req As Boolean, Optional VRul$, Optional VTxt$) As DAO.Field2
Set FdzTxt = NewFd(F, dbText, Req, TxtSz, ZLen, Expr, Dft, VRul, VTxt)
End Function

Function NewFdTY(F) As DAO.Field2
Set NewFdTY = NewFd(F, dbText, True, 20, ZLen:=False)
End Function

Private Sub ZZ()
Dim A As Variant
Dim B As DAO.DataTypeEnum
Dim C As Boolean
Dim D As Byte
Dim E$
NewFd A, B, C, D, C, E, E, E, E
FdzCrtDte A
FdzCur A
FdzDte A
FdzEleNmFld A, A
FdzFk A
FdzId A
FdzNm A
FdzPk A
StdFd A, E
FdzFdStr A
FdzTxt A, D, C, E, E, C, E, E
NewFdTY A
End Sub

Private Sub Z()
End Sub

