Attribute VB_Name = "MDao_Def_Td"
Option Explicit
Function NewTdTblFdAy(T, FdAy() As DAO.Field2) As DAO.TableDef
Dim O As New DAO.TableDef
O.Name = T
TdAppFdAy O, FdAy
Set NewTdTblFdAy = O
End Function
Function CvTd(A) As DAO.TableDef
Set CvTd = A
End Function

Sub TdAppFdAy(A As DAO.TableDef, FdAy() As DAO.Field2)
Dim I
For Each I In FdAy
    A.Fields.Append I
Next
End Sub

Sub TdAppIdFld(A As DAO.TableDef)
A.Fields.Append NewFd(A.Name)
End Sub

Sub TdAppLngFld(A As DAO.TableDef, FF)
TdAppFdAy A, ZFdAy(FF, dbLong)
End Sub

Sub TdAppLngTxt(A As DAO.TableDef, FF)
TdAppFdAy A, ZFdAy(FF, dbText)
End Sub

Sub TdAppTimStampFld(A As DAO.TableDef, F$)
A.Fields.Append NewFd(F, DAO.dbDate, Dft:="Now")
End Sub

Sub TdAddTxtFld(A As DAO.TableDef, FF0, Optional Req As Boolean, Optional Sz As Byte = 255)
Dim F
For Each F In CvNy(FF0)
    A.Fields.Append NewFd(F, dbText, Req, Sz)
Next
End Sub

Function TdFdScly(A As DAO.TableDef) As String()
Dim N$
N = A.Name & ";"
TdFdScly = AyAddPfx(SyzItrMap(A.Fields, "FdScl"), N)
End Function

Function TdFny(A As DAO.TableDef) As String()
TdFny = FnyzFds(A.Fields)
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

Sub ThwNETd(A As DAO.TableDef, B As DAO.TableDef)
Dim A1$: A1 = TdStrLines(A)
Dim B1$: B1 = TdStrLines(B)
If A1 <> B1 Then Stop
End Sub

Function SclzTd$(A As DAO.TableDef)
SclzTd = ApScl(A.Name, AddLib(A.OpenRecordset.RecordCount, "NRec"), AddLib(A.DateCreated, "CrtDte"), AddLib(A.LastUpdated, "UpdDte"))
End Function

Function SclzTdLy(A As DAO.TableDef) As String()
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

Function TdTyStr$(A As DAO.TableDefAttributeEnum)
TdTyStr = A
End Function

Private Function ZFdAy(FF, T As DAO.DataTypeEnum) As DAO.Field2()
Dim F
For Each F In CvNy(FF)
    PushObj ZFdAy, NewFd(F, T)
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

Function IsSysTd(A As DAO.TableDef) As Boolean
IsSysTd = A.Attributes And DAO.TableDefAttributeEnum.dbSystemObject <> 0
End Function

Function IsHidTd(A As DAO.TableDef) As Boolean
IsHidTd = A.Attributes And DAO.TableDefAttributeEnum.dbHiddenObject <> 0
End Function

