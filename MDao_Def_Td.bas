Attribute VB_Name = "MDao_Def_Td"
Option Explicit

Function CvTd(A) As DAO.TableDef
Set CvTd = A
End Function

Sub AddFdy(A As TableDef, Fdy() As DAO.Field2)
Dim I
For Each I In Fdy
    A.Fields.Append I
Next
End Sub

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

Function FnyzTd(A As DAO.TableDef) As String()
FnyzTd = Itn(A.Fields)
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

Sub ThwIfNETd(A As DAO.TableDef, B As DAO.TableDef)
Dim A1$(): A1 = TdLy(A)
Dim B1$(): B1 = TdLy(B)
If Not IsEqAy(A, B) Then Thw CSub, "Two 2 Td as diff", "Td-A Td-B", TdLy(A), TdLy(B)
End Sub

Sub DmpTdAy(TdAy() As DAO.TableDef)
Dim I
For Each I In TdAy
    D "------------------------"
    D TdLy(I)
Next
End Sub
Function TdLyzDb(A As Database) As String()
Dim T
For Each T In Tni(A)
    PushIAy TdLyzDb, TdLy(A.TableDefs(T))
Next
End Function

Function TdLyzT(A As Database, T$) As String()
TdLyzT = TdLy(A.TableDefs(T))
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

Private Function Fdy(FF$, T As DAO.DataTypeEnum) As DAO.Field2()
Dim I, F$
For Each I In TermAy(FF)
    F = I
    PushObj Fdy, Fd(F, T)
Next
End Function

Private Sub ZZ()
Dim A As Variant
Dim B As DAO.TableDef
Dim C() As DAO.Field2
Dim D$
Dim E As Boolean
Dim F As Byte
Dim G As DAO.TableDefAttributeEnum
CvTd A
AddFldzId B
End Sub

Function IsSysTd(A As DAO.TableDef) As Boolean
IsSysTd = A.Attributes And DAO.TableDefAttributeEnum.dbSystemObject <> 0
End Function

Function IsHidTd(A As DAO.TableDef) As Boolean
IsHidTd = A.Attributes And DAO.TableDefAttributeEnum.dbHiddenObject <> 0
End Function

