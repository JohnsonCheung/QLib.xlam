Attribute VB_Name = "MDao_Def_Td"
Option Explicit
Function NewTdTblFdAy(T, FdAy() As Dao.Field2) As Dao.TableDef
Dim O As New Dao.TableDef
O.Name = T
TdAppFdAy O, FdAy
Set NewTdTblFdAy = O
End Function
Function CvTd(A) As Dao.TableDef
Set CvTd = A
End Function

Sub TdAppFdAy(A As Dao.TableDef, FdAy() As Dao.Field2)
Dim I
For Each I In FdAy
    A.Fields.Append I
Next
End Sub

Sub TdAppIdFld(A As Dao.TableDef)
A.Fields.Append FdzId(A.Name)
End Sub

Sub TdAppLngFld(A As Dao.TableDef, FF)
TdAppFdAy A, ZFdAy(FF, dbLong)
End Sub

Sub TdAppLngTxt(A As Dao.TableDef, FF)
TdAppFdAy A, ZFdAy(FF, dbText)
End Sub

Sub TdAppTimStampFld(A As Dao.TableDef, F$)
A.Fields.Append Fd(F, Dao.dbDate, Dft:="Now")
End Sub

Sub TdAddTxtFld(A As Dao.TableDef, FF0, Optional Req As Boolean, Optional Sz As Byte = 255)
Dim F
For Each F In CvNy(FF0)
    A.Fields.Append Fd(F, dbText, Req, Sz)
Next
End Sub

Function TdFdScly(A As Dao.TableDef) As String()
Dim N$
N = A.Name & ";"
TdFdScly = AyAddPfx(SyzItrMap(A.Fields, "FdScl"), N)
End Function

Function TdFny(A As Dao.TableDef) As String()
TdFny = FnyzFds(A.Fields)
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

Sub ThwNETd(A As Dao.TableDef, B As Dao.TableDef)
Dim A1$(): A1 = TdFdLy(A)
Dim B1$(): B1 = TdFdLy(B)
If Not IsEqAy(A, B) Then Thw CSub, "Two 2 Td as diff", "Td-A Td-B", TdFdLy(A), TdFdLy(B)
End Sub
Function TdFdLy(A As Dao.TableDef) As String()
Dim O$()
PushI O, TdStr(A)
Dim F As Dao.Field
For Each F In A.Fields
    PushI O, FdStr(F)
Next
TdFdLy = O
End Function
Function SclzTd$(A As Dao.TableDef)
SclzTd = ApScl(A.Name, AddLib(A.OpenRecordset.RecordCount, "NRec"), AddLib(A.DateCreated, "CrtDte"), AddLib(A.LastUpdated, "UpdDte"))
End Function

Function SclzTdLy(A As Dao.TableDef) As String()
SclzTdLy = AyAdd(Sy(SclzTd(A)), TdFdScly(A))
End Function

Function SclzTdLy_AddPfx(TdLy$()) As String()
Dim O$(), U&, J&, X
U = UB(TdLy)
If U = -1 Then Exit Function
ReDim O(U)
For Each X In Itr(TdLy)
    O(J) = IIf(J = 0, "Td;", "Fd;") & X
    J = J + 1
Next
SclzTdLy_AddPfx = O
End Function

Function TdTyStr$(A As Dao.TableDefAttributeEnum)
TdTyStr = A
End Function

Private Function ZFdAy(FF, T As Dao.DataTypeEnum) As Dao.Field2()
Dim F
For Each F In CvNy(FF)
    PushObj ZFdAy, Fd(F, T)
Next
End Function

Private Sub ZZ()
Dim A As Variant
Dim B As Dao.TableDef
Dim C() As Dao.Field2
Dim D$
Dim E As Boolean
Dim F As Byte
Dim G As Dao.TableDefAttributeEnum
CvTd A
TdAppFdAy B, C
TdAppIdFld B
TdAppLngFld B, A
TdAppLngTxt B, A
TdAppTimStampFld B, D
TdAddTxtFld B, A, E, F
TdFdScly B
TdFny B
IsEqTd B, B
ThwNETd B, B
SclzTd B
SclzTdLy B
End Sub

Private Sub Z()
End Sub

Function IsSysTd(A As Dao.TableDef) As Boolean
IsSysTd = A.Attributes And Dao.TableDefAttributeEnum.dbSystemObject <> 0
End Function

Function IsHidTd(A As Dao.TableDef) As Boolean
IsHidTd = A.Attributes And Dao.TableDefAttributeEnum.dbHiddenObject <> 0
End Function

