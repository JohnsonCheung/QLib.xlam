Attribute VB_Name = "QDao_Def_Td"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Def_Td."
Private Const Asm$ = "QDao"

Property Get TmpTd() As Dao.TableDef
Dim Fdy() As Dao.Field2
PushObj Fdy, FdzTxt("F1")
Set TmpTd = TdzTF("Tmp", Fdy)
End Property

Function CvTd(A) As Dao.TableDef
Set CvTd = A
End Function

Function TdzTF(T, Fdy() As Dao.Field2) As Dao.TableDef
Dim O As New TableDef
O.Name = T
AddFdy O, Fdy
Set TdzTF = O
End Function

Sub AddFdy(A As TableDef, Fdy() As Dao.Field2)
Dim I: For Each I In Fdy
    A.Fields.Append I
Next
End Sub

Sub AddFldzId(A As Dao.TableDef)
A.Fields.Append FdzId(A.Name)
End Sub

Sub AddFldzLng(A As Dao.TableDef, FF$)
AddFdy A, Fdy(FF, dbLong)
End Sub

Sub AddFldzTimstmp(A As Dao.TableDef, F$)
A.Fields.Append Fd(F, Dao.dbDate, Dft:="Now")
End Sub

Sub AddFldzTxt(A As Dao.TableDef, FF$, Optional Req As Boolean, Optional Si As Byte = 255)
Dim F$, I
For Each I In TermAy(FF)
    F = I
    A.Fields.Append Fd(F, dbText, Req, Si)
Next
End Sub

Function FnyzTd(A As Dao.TableDef) As String()
FnyzTd = Itn(A.Fields)
End Function

Function IsEqTd(A As Dao.TableDef, B As Dao.TableDef) As Boolean
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

Sub ThwIf_NETd(A As Dao.TableDef, B As Dao.TableDef)
Dim A1$(): A1 = TdLy(A)
Dim B1$(): B1 = TdLy(B)
If Not IsEqAy(A, B) Then Thw CSub, "Two 2 Td as diff", "Td-A Td-B", TdLy(A), TdLy(B)
End Sub

Sub DmpTdAy(TdAy() As Dao.TableDef)
Dim I
For Each I In TdAy
    D "------------------------"
    D TdLy(I)
Next
End Sub
Function TdLyzDb(D As Database) As String()
Dim T
For Each T In Tni(D)
    PushIAy TdLyzDb, TdLy(D.TableDefs(T))
Next
End Function

Function TdLyzT(D As Database, T) As String()
TdLyzT = TdLy(D.TableDefs(T))
End Function

Function TdLy(Td) As String()
Dim O$(), A As Dao.TableDef
Set A = Td
PushI TdLy, TdStr(A)
Dim F As Dao.Field
For Each F In A.Fields
    PushI TdLy, FdStr(F)
Next
End Function

Private Function Fdy(FF$, T As Dao.DataTypeEnum) As Dao.Field2()
Dim I, F$
For Each I In TermAy(FF)
    F = I
    PushObj Fdy, Fd(F, T)
Next
End Function

Private Sub Z()
Dim A As Variant
Dim B As Dao.TableDef
Dim C() As Dao.Field2
Dim D$
Dim E As Boolean
Dim F As Byte
Dim G As Dao.TableDefAttributeEnum
CvTd A
AddFldzId B
End Sub

Function IsTdSys(A As Dao.TableDef) As Boolean
IsTdSys = A.Attributes And Dao.TableDefAttributeEnum.dbSystemObject <> 0
End Function

Function IsTdHid(A As Dao.TableDef) As Boolean
IsTdHid = A.Attributes And Dao.TableDefAttributeEnum.dbHiddenObject <> 0
End Function



Function TdStr$(A As Dao.TableDef)
Dim T$, Id$, S$, R$
    T = A.Name
    If HasStdPkzTd(A) Then Id = "*Id"
    Dim Pk$(): Pk = Sy(T & "Id")
    Dim Sk$(): Sk = SkFnyzTd(A)
    If HasStdSkzTd(A) Then S = TLin(RplAy(Sk, T, "*")) & " |"
    R = TLin(CvSy(MinusAyAp(FnyzTd(A), Pk, Sk)))
TdStr = JnSpc(SyNB(T, Id, S, R))
End Function

Function FnyzTdLy(TdLy$()) As String()
Dim O$(), TdStr$, I
For Each I In Itr(TdLy)
    TdStr = I
    PushIAy O, FnyzTdLin(TdStr)
Next
FnyzTdLy = CvSy(AwDist(O))
End Function

Function TdStrzT$(D As Database, T)
TdStrzT = TdStr(D.TableDefs(T))
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

