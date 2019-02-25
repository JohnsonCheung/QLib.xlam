Attribute VB_Name = "MDao_Prp"
Option Explicit

Property Get FldDesz$(A As Database, T, F)
FldDesz = FldPrpz(A, T, F, C_Des)
End Property
Property Let FldDesz(A As Database, T, F, Des$)
FldPrpz(A, T, F, C_Des) = Des
End Property


Function SampFd() As DAO.Field2
Set SampFd = SampDb_DutyDta.TableDefs("Permit").Fields("Permit")
End Function
Private Sub Z_PrpNy()
D PrpNyFd(SampFd)
'D PrpNyFd(SampDb_DutyDta.TableDefs("Permit").Fields("Permit"))
End Sub
Function PrpNyFd(A As DAO.Field) As String()
PrpNyFd = Itn(A.Properties)
End Function
Private Sub Z_PrpDryFd()
'DmpDry PrpDryFd(SampDb_DutyDta.TableDefs("Permit").Fields("Permit"))
DmpDry PrpDryFd(SampFd)
End Sub
Function PrpDryFd(A As DAO.Field) As Variant()
Dim Prp
For Each Prp In Itn(A.Properties)
    PushI PrpDryFd, Array(Prp, PrpPrpsP(A.Properties, Prp))
Next
End Function
Function PrpPrpsP(A As DAO.Properties, P)
On Error GoTo X
PrpPrpsP = A(P).Value
X:
End Function

Property Get PrpTFP(T, F, P)
PrpTFP = FldPrpz(CDb, T, F, P)
End Property

Property Let PrpTFP(T, F, P, V)
FldPrpz(CDb, T, F, P) = V
End Property

Private Sub Z_PrpTFP()
'XRfh_TmpTbl
Dim P$
P = "Ele"
Ept = 123
GoSub Tst
Exit Sub
Tst:
    PrpTFP("Tmp", "F1", P) = Ept
'    Act = PrpTFP("Tmp", "F1", P)
    C
    Return
End Sub

Function DesFd$(A As DAO.Field)
DesFd = PrpPrpsP(A.Properties, C_Des)
End Function


Private Sub Z()
MDao_Z_Prp_Fld:
End Sub
Function Any_Prps_P(Prps As DAO.Properties, P) As Boolean
Any_Prps_P = HasItn(Prps, P)
End Function

Function PrpVal(A As DAO.Properties, PrpNm$)
On Error Resume Next
PrpVal = A(PrpNm).Value
End Function
Private Sub ZZ_TblPrpz()
Dim D As Database: Set D = TmpDb
Drpz D, "Tmp"
RunQz D, "Create Table Tmp (F1 Text)"
TblPrpz(D, "Tmp", "XX") = "AFdf"
Debug.Assert TblPrpz(CDb, "Tmp", "XX") = "AFdf"
End Sub

Function CrtPrp(A As Database, T, P$, V) As DAO.Property
Set CrtPrp = A.TableDefs(T).CreateProperty(P, DaoTyVal(V), V) ' will break if V=""
End Function

Function HasDbtPrp(A As Database, T, P) As Boolean
HasDbtPrp = HasItn(A.TableDefs(T).Properties, P)
End Function

Property Get TblPrp(T, P)
TblPrp = TblPrpz(CDb, T, P)
End Property

Property Let TblPrp(T, P, V)
TblPrpz(CDb, T, P) = V
End Property

Property Get TblPrpz(A As Database, T, P)
If Not HasDbtPrp(A, T, P) Then Exit Property
TblPrpz = A.TableDefs(T).Properties(P).Value
End Property

Property Let TblPrpz(A As Database, T, P, V)
Dim Td As DAO.TableDef: Set Td = A.TableDefs(T)
SetDaoPrp Td, Td.Properties, P, V
End Property

Sub SetDaoPrp(DaoObj, Prps As DAO.Properties, P, V)
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
Dim Fd As DAO.Field: Set Fd = A.TableDefs(T).Fields(F)
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
TblPrpz(CDb, T, P) = V
End Property
