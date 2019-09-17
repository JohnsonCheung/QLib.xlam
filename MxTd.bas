Attribute VB_Name = "MxTd"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxTd."

Sub AddFldzId(A As DAO.TableDef)
A.Fields.Append FdzId(A.Name)
End Sub

Sub AddFldzLng(A As DAO.TableDef, FF$)
AddFdy A, Fdy(FF, dbLong)
End Sub

Sub AddFldzTimstmp(A As DAO.TableDef, F$)
A.Fields.Append Fd(F, DAO.dbDate, Dft:="Now")
End Sub

Sub AddFldzTxt(A As DAO.TableDef, FF$, Optional Req As Boolean, Optional Si As Byte = 255)
Dim F$, I
For Each I In TermAy(FF)
    F = I
    A.Fields.Append Fd(F, dbText, Req, Si)
Next
End Sub

Function CvTd(A) As DAO.TableDef
Set CvTd = A
End Function

Sub DmpTdAy(TdAy() As DAO.TableDef)
Dim I
For Each I In TdAy
    D "------------------------"
    D TdLy(I)
Next
End Sub

Function Fdy(FF$, T As DAO.DataTypeEnum) As DAO.Field2()
Dim I, F$
For Each I In TermAy(FF)
    F = I
    PushObj Fdy, Fd(F, T)
Next
End Function

Function FnyzTd(A As DAO.TableDef) As String()
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

Function IsEqTd(A As DAO.TableDef, B As DAO.TableDef) As Boolean
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

Function IsTdHid(A As DAO.TableDef) As Boolean
IsTdHid = A.Attributes And DAO.TableDefAttributeEnum.dbHiddenObject <> 0
End Function

Function IsTdSys(A As DAO.TableDef) As Boolean
IsTdSys = A.Attributes And DAO.TableDefAttributeEnum.dbSystemObject <> 0
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
Dim O$(), A As DAO.TableDef
Set A = Td
PushI TdLy, TdStr(A)
Dim F As DAO.Field
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

Function TdStr$(A As DAO.TableDef)
Dim T$, Id$, S$, R$
    T = A.Name
    If HasStdPkzTd(A) Then Id = "*Id"
    Dim Pk$(): Pk = Sy(T & "Id")
    Dim Sk$(): Sk = SkFnyzTd(A)
    If HasStdSkzTd(A) Then S = TLin(AmRpl(Sk, T, "*")) & " |"
    R = TLin(CvSy(AyMinusAp(FnyzTd(A), Pk, Sk)))
TdStr = JnSpc(SyNB(T, Id, S, R))
End Function

Function TdStrzT$(D As Database, T)
TdStrzT = TdStr(D.TableDefs(T))
End Function

Function TdzTF(T, Fdy() As DAO.Field2) As DAO.TableDef
Dim O As New TableDef
O.Name = T
AddFdy O, Fdy
Set TdzTF = O
End Function

Sub ThwIf_NETd(A As DAO.TableDef, B As DAO.TableDef)
Dim A1$(): A1 = TdLy(A)
Dim B1$(): B1 = TdLy(B)
If Not IsEqAy(A, B) Then Thw CSub, "Two 2 Td as diff", "Td-A Td-B", TdLy(A), TdLy(B)
End Sub

Property Get TmpTd() As DAO.TableDef
Dim Fdy() As DAO.Field2
PushObj Fdy, FdzTxt("F1")
Set TmpTd = TdzTF("Tmp", Fdy)
End Property


Sub AddPk(A As DAO.TableDef)
'Any Pk Fields in A.Fields?, if no exit sub
Dim F As DAO.Field2, IdFldNm$, J%
IdFldNm = A.Name & "Id"
If IsFdId(A.Fields(0), A.Name) Then
    A.Indexes.Append PkizT(A.Name)
    Exit Sub
End If
For J = 2 To A.Fields.Count
    If A.Fields(J).Name = IdFldNm Then Thw CSub, "The Table Id fields must be the fst fld", "I-th", J
Next
End Sub

Sub AddSk(A As DAO.TableDef, Skff$)
Dim SkFny$(): SkFny = TermAy(Skff): If Si(SkFny) = 0 Then Exit Sub
A.Indexes.Append NewSkIdx(A, SkFny)
End Sub

Function CvIdxFds(A) As DAO.IndexFields
Set CvIdxFds = A
End Function

Function IsFdId(A As DAO.Field2, T) As Boolean
If A.Name <> T & "Id" Then Exit Function
If A.Attributes <> DAO.FieldAttributeEnum.dbAutoIncrField Then Exit Function
If A.Type <> dbLong Then Exit Function
IsFdId = True
End Function

Function NewSkIdx(T As DAO.TableDef, SkFny$()) As DAO.Index
Const CSub$ = CMod & "NewSkIdx"
Dim O As New DAO.Index
O.Name = "SecondaryKey"
O.Unique = True
If Not HasEleAy(FnyzTd(T), SkFny) Then
    Thw CSub, "Given Td does not contain all given-SkFny", "Missing-SkFny Td-Name Td-Fny Given-SkFny", T.Name & "Id", AyMinus(SkFny, FnyzTd(T)), T.Name, FnyzTd(T), SkFny
End If
Dim IdxFds As DAO.IndexFields, I
Set IdxFds = CvIdxFds(O.Fields)
For Each I In SkFny
    IdxFds.Append Fd(CStr(I))
Next
Set NewSkIdx = O
End Function

Function PkizT(T) As DAO.Index
Dim O As New DAO.Index
O.Name = "PrimaryKey"
O.Primary = True
CvIdxFds(O.Fields).Append FdzId(T & "Id")
Set PkizT = O
End Function

Function NewTd(T) As DAO.TableDef
Set NewTd = New DAO.TableDef
NewTd.Name = T
End Function
