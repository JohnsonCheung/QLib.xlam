Attribute VB_Name = "MxTd"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxTd."

Sub AddFldzId(A As dao.TableDef)
A.Fields.Append FdzId(A.Name)
End Sub

Sub AddFldzLng(A As dao.TableDef, FF$)
AddFdy A, Fdy(FF, dbLong)
End Sub

Sub AddFldzTimstmp(A As dao.TableDef, F$)
A.Fields.Append Fd(F, dao.dbDate, Dft:="Now")
End Sub

Sub AddFldzTxt(A As dao.TableDef, FF$, Optional Req As Boolean, Optional Si As Byte = 255)
Dim F$, I
For Each I In TermAy(FF)
    F = I
    A.Fields.Append Fd(F, dbText, Req, Si)
Next
End Sub

Function CvTd(A) As dao.TableDef
Set CvTd = A
End Function

Sub DmpTdAy(TdAy() As dao.TableDef)
Dim I
For Each I In TdAy
    D "------------------------"
    D TdLy(I)
Next
End Sub

Private Function Fdy(FF$, T As dao.DataTypeEnum) As dao.Field2()
Dim I, F$
For Each I In TermAy(FF)
    F = I
    PushObj Fdy, Fd(F, T)
Next
End Function

Function FnyzTd(A As dao.TableDef) As String()
FnyzTd = Itn(A.Fields)
End Function

Function FnyzTdLy(TdLy$()) As String()
Dim O$(), TdStr$, I
For Each I In Itr(TdLy)
    TdStr = I
    PushIAy O, FnyzTdLin(TdStr)
Next
FnyzTdLy = CvSy(AwDist(O))
End Function

Function IsEqTd(A As dao.TableDef, B As dao.TableDef) As Boolean
With A
Select Case True
Case .Name <> B.Name
Case .Attributes <> B.Attributes
Case Not IsEqIdxs(.Indexes, B.Indexes)
'Case Not FdsIsEq(.Fields, B.Fields)
Case Else: IsEqTd = True
End Select
End With
End Function

Function IsTdHid(A As dao.TableDef) As Boolean
IsTdHid = A.Attributes And dao.TableDefAttributeEnum.dbHiddenObject <> 0
End Function

Function IsTdSys(A As dao.TableDef) As Boolean
IsTdSys = A.Attributes And dao.TableDefAttributeEnum.dbSystemObject <> 0
End Function

Function SkFnyzTdLin(TdLin) As String()
Dim A1$, T$, Rst$
    A1 = Bef(TdLin, "|")
    If A1 = "" Then Exit Function
AsgTRst A1, T, Rst
T = RmvSfx(T, "*")
Rst = Replace(Rst, "*", T)
SkFnyzTdLin = SyzSS(Rst)
End Function

Function TdLy(Td) As String()
Dim O$(), A As dao.TableDef
Set A = Td
PushI TdLy, TdStr(A)
Dim F As dao.Field
For Each F In A.Fields
    PushI TdLy, FdStr(F)
Next
End Function

Function TdLyzDb(D As Database) As String()
Dim T
For Each T In Tni(D)
    PushIAy TdLyzDb, TdLy(D.TableDefs(T))
Next
End Function

Function TdLyzT(D As Database, T) As String()
TdLyzT = TdLy(D.TableDefs(T))
End Function

Function TdStr$(A As dao.TableDef)
Dim T$, Id$, S$, R$
    T = A.Name
    If HasStdPkzTd(A) Then Id = "*Id"
    Dim Pk$(): Pk = Sy(T & "Id")
    Dim Sk$(): Sk = SkFnyzTd(A)
    If HasStdSkzTd(A) Then S = TLin(RplAy(Sk, T, "*")) & " |"
    R = TLin(CvSy(MinusAyAp(FnyzTd(A), Pk, Sk)))
TdStr = JnSpc(SyNB(T, Id, S, R))
End Function

Function TdStrzT$(D As Database, T)
TdStrzT = TdStr(D.TableDefs(T))
End Function

Function TdzTF(T, Fdy() As dao.Field2) As dao.TableDef
Dim O As New TableDef
O.Name = T
AddFdy O, Fdy
Set TdzTF = O
End Function

Sub ThwIf_NETd(A As dao.TableDef, B As dao.TableDef)
Dim A1$(): A1 = TdLy(A)
Dim B1$(): B1 = TdLy(B)
If Not IsEqAy(A, B) Then Thw CSub, "Two 2 Td as diff", "Td-A Td-B", TdLy(A), TdLy(B)
End Sub

Property Get TmpTd() As dao.TableDef
Dim Fdy() As dao.Field2
PushObj Fdy, FdzTxt("F1")
Set TmpTd = TdzTF("Tmp", Fdy)
End Property

Private Sub Z()
Dim A As Variant
Dim B As dao.TableDef
Dim C() As dao.Field2
Dim D$
Dim E As Boolean
Dim F As Byte
Dim G As dao.TableDefAttributeEnum
CvTd A
AddFldzId B
End Sub

Private Sub AddPk(A As dao.TableDef)
'Any Pk Fields in A.Fields?, if no exit sub
Dim F As dao.Field2, IdFldNm$, J%
IdFldNm = A.Name & "Id"
If IsFdId(A.Fields(0), A.Name) Then
    A.Indexes.Append PkizT(A.Name)
    Exit Sub
End If
For J = 2 To A.Fields.Count
    If A.Fields(J).Name = IdFldNm Then Thw CSub, "The Table Id fields must be the fst fld", "I-th", J
Next
End Sub

Private Sub AddSk(A As dao.TableDef, Skff$)
Dim SkFny$(): SkFny = TermAy(Skff): If Si(SkFny) = 0 Then Exit Sub
A.Indexes.Append NewSkIdx(A, SkFny)
End Sub

Private Function CvIdxFds(A) As dao.IndexFields
Set CvIdxFds = A
End Function

Private Function IsFdId(A As dao.Field2, T) As Boolean
If A.Name <> T & "Id" Then Exit Function
If A.Attributes <> dao.FieldAttributeEnum.dbAutoIncrField Then Exit Function
If A.Type <> dbLong Then Exit Function
IsFdId = True
End Function

Function NewSkIdx(T As dao.TableDef, SkFny$()) As dao.Index
Const CSub$ = CMod & "NewSkIdx"
Dim O As New dao.Index
O.Name = "SecondaryKey"
O.Unique = True
If Not HasEleAy(FnyzTd(T), SkFny) Then
    Thw CSub, "Given Td does not contain all given-SkFny", "Missing-SkFny Td-Name Td-Fny Given-SkFny", T.Name & "Id", MinusAy(SkFny, FnyzTd(T)), T.Name, FnyzTd(T), SkFny
End If
Dim IdxFds As dao.IndexFields, I
Set IdxFds = CvIdxFds(O.Fields)
For Each I In SkFny
    IdxFds.Append Fd(CStr(I))
Next
Set NewSkIdx = O
End Function

Private Function PkizT(T) As dao.Index
Dim O As New dao.Index
O.Name = "PrimaryKey"
O.Primary = True
CvIdxFds(O.Fields).Append FdzId(T & "Id")
Set PkizT = O
End Function

Function TdzNm(T) As dao.TableDef
Set TdzNm = New TableDef
TdzNm.Name = T
End Function