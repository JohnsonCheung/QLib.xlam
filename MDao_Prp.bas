Attribute VB_Name = "MDao_Prp"
Option Explicit

Property Get FldDes$(A As Database, T$, F$)
FldDes = FldPrp(A, T, F, C_Des)
End Property

Property Let FldDes(A As Database, T$, F$, Des$)
FldPrp(A, T, F, C_Des) = Des
End Property

Private Sub Z_PrpNy()
Dim Db As Database: Set Db = SampDbzDutyDta
Dim Fd As DAO.Field
Set Fd = FdzTF(Db, "Permit", "Permit")
D PrpNyzFd(Fd)
End Sub

Function PrpNyzFd(A As DAO.Field) As String()
PrpNyzFd = Itn(A.Properties)
End Function

Private Sub Z_PrpDryzFd()
Dim Db As Database: Set Db = SampDbzDutyDta
Dim Fd As DAO.Field
Dim Rs As DAO.Recordset
Set Rs = RszT(Db, "Permit")
Set Fd = Rs.Fields("Permit")
Debug.Print Fd.Value
DmpDry PrpDryzFd(Fd)
End Sub

Function PrpDryzFd(A As DAO.Field) As Variant()
Dim PrpV, I, P$, V
For Each I In Itn(A.Properties)
    P = I
    V = PrpVal(A, P)
    PushI PrpDryzFd, Array(P, V, TypeName(V))
Next
End Function

Property Let PrpVal(WiDaoPrps As Object, P$, V)
Dim Prps As DAO.Properties
Set Prps = WiDaoPrps.Properties
If HasItn(Prps, P) Then
    Prps(P).Value = V
Else
    Prps.Append WiDaoPrps.CreateProperty(P, DaoTyzVal(V), V) ' will break if V=""
End If
End Property

Property Get PrpVal(WiDaoPrps As Object, P$)
Dim Prps As DAO.Properties
Set Prps = WiDaoPrps.Properties
On Error GoTo X
PrpVal = Prps(P).Value
X:
End Property

Private Sub Z_FldPrp()
Dim P$, Db As Database, T$, F$, V
GoSub T0
Exit Sub
T0:
    Set Db = TmpDb
    RunQ Db, "Create Table Tmp (AA Text)"
    T = "Tmp"
    F = "AA"
    P = "Ele"
    V = "Ele1234"
    GoTo Tst
Tst:
    FldPrp(Db, T, F, P) = V
    Ass FldPrp(Db, T, F, P) = V
    Dim Fd As DAO.Field: Set Fd = FdzTF(Db, T, F)
    Stop
    DmpDry PrpDryzFd(Fd)
    Return
End Sub

Property Get FldDeszTd$(A As DAO.Field)
FldDeszTd = PrpVal(A.Properties, C_Des)
End Property

Property Let FldDeszTd(A As DAO.Field, Des$)

End Property

Private Sub Z()
MDao_Z_Prp_Fld:
End Sub

Private Sub ZZ_TblPrp()
Dim D As Database: Set D = TmpDb
DrpT D, "Tmp"
RunQ D, "Create Table Tmp (F1 Text)"
TblPrp(D, "Tmp", "XX") = "AFdf"
Debug.Assert TblPrp(D, "Tmp", "XX") = "AFdf"
End Sub

Function HasDbtPrp(A As Database, T, P$) As Boolean
HasDbtPrp = HasItn(A.TableDefs(T).Properties, P)
End Function

Property Get TblPrp(A As Database, T$, P$)
If Not HasDbtPrp(A, T, P) Then Exit Property
TblPrp = A.TableDefs(T).Properties(P).Value
End Property

Property Let TblPrp(A As Database, T$, P$, V)
Dim Td As DAO.TableDef: Set Td = A.TableDefs(T)
SetDaoPrp Td, Td.Properties, P, V
End Property

Sub SetDaoPrp(WiDaoPrps As Object, Prps As DAO.Properties, P$, V)
If HasItn(Prps, P) Then
    If IsEmp(V) Then
        Prps.Delete P
        Exit Sub
    End If
    Prps(P).Value = V
Else
    Prps.Append WiDaoPrps.CreateProperty(P, DaoTyzVal(V), V)
End If
End Sub

Property Let FldPrp(A As Database, T$, F$, P$, V)
Dim Fd As DAO.Field: Set Fd = A.TableDefs(T).Fields(F)
SetDaoPrp Fd, Fd.Properties, P, V
End Property

Property Get FldPrp(A As Database, T$, F$, P$)
If Not HasFldPrp(A, T, F, P) Then Exit Property
FldPrp = A.TableDefs(T).Fields(F).Properties(P).Value
End Property

Function HasFldPrp(A As Database, T$, F$, P$) As Boolean
HasFldPrp = HasItn(A.TableDefs(T).Fields(F).Properties, P)
End Function

