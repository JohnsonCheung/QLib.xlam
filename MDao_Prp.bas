Attribute VB_Name = "MDao_Prp"
Option Explicit

Property Get FldDes$(A As Database, T, F)
FldDes = FldPrpz(A, T, F, C_Des)
End Property
Property Let FldDes(A As Database, T, F, Des$)
FldPrpz(A, T, F, C_Des) = Des
End Property


Function SampFd() As Dao.Field2
Set SampFd = SampDb_DutyDta.TableDefs("Permit").Fields("Permit")
End Function
Private Sub Z_PrpNy()
D PrpNyFd(SampFd)
'D PrpNyFd(SampDb_DutyDta.TableDefs("Permit").Fields("Permit"))
End Sub
Function PrpNyFd(A As Dao.Field) As String()
PrpNyFd = Itn(A.Properties)
End Function
Private Sub Z_PrpDryzFd()
'DmpDry PrpDryzFd(SampDb_DutyDta.TableDefs("Permit").Fields("Permit"))
DmpDry PrpDryzFd(SampFd)
End Sub
Function PrpDryzFd(A As Dao.Field) As Variant()
Dim Prp
For Each Prp In Itn(A.Properties)
    PushI PrpDryzFd, Array(Prp, ValzDaoPrp(A.Properties, Prp))
Next
End Function

Function ValzDaoPrp(A As Dao.Properties, P)
On Error GoTo X
ValzDaoPrp = A(P).Value
X:
End Function
Property Get FldPrp(A As Database, T, F, P)

End Property
Property Let FldPrp(A As Database, T, F, P, V)

End Property
Private Sub Z_FldPrp()
'XRfh_TmpTbl
Dim P$, Db As Database
P = "Ele"
Ept = 123
GoSub Tst
Exit Sub
Tst:
    FldPrp(Db, "Tmp", "F1", P) = Ept
'    Act = PrpTFP("Tmp", "F1", P)
    C
    Return
End Sub

Function DeszFd$(A As Dao.Field)
DeszFd = ValzDaoPrp(A.Properties, C_Des)
End Function


Private Sub Z()
MDao_Z_Prp_Fld:
End Sub
Function Any_Prps_P(Prps As Dao.Properties, P) As Boolean
Any_Prps_P = HasItn(Prps, P)
End Function

Function PrpVal(A As Dao.Properties, PrpNm$)
On Error Resume Next
PrpVal = A(PrpNm).Value
End Function
Private Sub ZZ_TblPrp()
Dim D As Database: Set D = TmpDb
DrpT D, "Tmp"
RunQ D, "Create Table Tmp (F1 Text)"
TblPrp(D, "Tmp", "XX") = "AFdf"
Debug.Assert TblPrp(CDb, "Tmp", "XX") = "AFdf"
End Sub

Function CrtPrp(A As Database, T, P$, V) As Dao.Property
Set CrtPrp = A.TableDefs(T).CreateProperty(P, DaoTyVal(V), V) ' will break if V=""
End Function

Function HasDbtPrp(A As Database, T, P) As Boolean
HasDbtPrp = HasItn(A.TableDefs(T).Properties, P)
End Function

Property Get TblPrp(A As Database, T, P)
If Not HasDbtPrp(A, T, P) Then Exit Property
TblPrp = A.TableDefs(T).Properties(P).Value
End Property

Property Let TblPrp(A As Database, T, P, V)
Dim Td As Dao.TableDef: Set Td = A.TableDefs(T)
SetDaoPrp Td, Td.Properties, P, V
End Property

Sub SetDaoPrp(DaoObj, Prps As Dao.Properties, P, V)
If HasItn(Prps, P) Then
    If IsEmp(V) Then
        Prps.Delete P
        Exit Sub
    End If
    Prps(P).Value = V
Else
    Prps.Append DaoObj.CreateProperty(P, DaoTyVal(V), V)
End If
End Sub
Property Let FldPrpz(A As Database, T, F, P, V)
Dim Fd As Dao.Field: Set Fd = A.TableDefs(T).Fields(F)
SetDaoPrp Fd, Fd.Properties, P, V
End Property

Property Get FldPrpz(A As Database, T, F, P)
If Not HasFldPrpz(A, T, F, P) Then Exit Property
FldPrpz = A.TableDefs(T).Fields(F).Properties(P).Value
End Property

Function HasFldPrpz(A As Database, T, F, P) As Boolean
HasFldPrpz = HasItn(A.TableDefs(T).Fields(F).Properties, P)
End Function

Property Let PrpTP(T, P$, V)
TblPrp(CDb, T, P) = V
End Property
