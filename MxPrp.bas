Attribute VB_Name = "MxPrp"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxPrp."

Function FldDes$(D As Database, T, F$)
FldDes = FldPrp(D, T, F, C_Des)
End Function

Function FldDeszTd$(A As dao.Field)
FldDeszTd = VzOPrps(A.Properties, C_Des)
End Function

Function FldPrp(D As Database, T, F$, P$)
If Not HasFldPrp(D, T, F, P) Then Exit Function
FldPrp = D.TableDefs(T).Fields(F).Properties(P).Value
End Function

Function HasDbtPrp(D As Database, T, P) As Boolean
HasDbtPrp = HasItn(D.TableDefs(T).Properties, P)
End Function

Function HasFldPrp(D As Database, T, F$, P$) As Boolean
HasFldPrp = HasItn(D.TableDefs(T).Fields(F).Properties, P)
End Function

Function PrpDyoFd(A As dao.Field) As Variant()
Dim PrpV, I, P$, V
For Each I In Itn(A.Properties)
    P = I
    V = VzOPrps(A, P)
    PushI PrpDyoFd, Array(P, V, TypeName(V))
Next
End Function

Function PrpNyzFd(A As dao.Field) As String()
PrpNyzFd = Itn(A.Properties)
End Function

Function PrpszO(ObjWiPrps) As dao.Properties
On Error GoTo X
Set PrpszO = ObjWiPrps.Properties
Exit Function
X:
    Dim E$: E = Err.Description
    Thw CSub, "Obj does not have prp-[Properties]", "Obj-TyNm Er", TypeName(ObjWiPrps), E
End Function

Sub SetDaoPrp(WiDaoPrps As Object, Prps As dao.Properties, P, V)
If HasItn(Prps, P) Then
    If IsEmp(V) Then
        Prps.Delete P
        Exit Sub
    End If
    Prps(P).Value = V
Else
    Prps.Append WiDaoPrps.CreateProperty(P, DaoTyzV(V), V)
End If
End Sub

Sub SetFldDes(D As Database, T, F$, Des$)
FldPrp(D, T, F, C_Des) = Des
End Sub

Sub SetFldDeszTd(A As dao.Field, Des$)

End Sub

Sub SetFldPrp(D As Database, T, F$, P$, V)
Dim Fd As dao.Field: Set Fd = D.TableDefs(T).Fields(F)
SetDaoPrp Fd, Fd.Properties, P, V
End Sub

Sub SetTblPrp(D As Database, T, P, V)
Dim Td As dao.TableDef: Set Td = D.TableDefs(T)
SetDaoPrp Td, Td.Properties, P, V
End Sub

Sub SetVzOPrps(ObjWiPrps, P$, V)
Dim Prps As dao.Properties: Set Prps = PrpszO(ObjWiPrps)
If HasItn(Prps, P) Then
    Prps(P).Value = V
Else
    Prps.Append ObjWiPrps.CreateProperty(P, DaoTyzV(V), V) ' will break if V=""
End If
End Sub

Function TblPrp(D As Database, T, P)
If Not HasDbtPrp(D, T, P) Then Exit Function
TblPrp = D.TableDefs(T).Properties(P).Value
End Function

Function VzOPrps(ObjWiPrps, P$)
'Ret : #Val-fm-ObjWithPrps ! Notes: Just passing @ObjWiPrps.Properties is Ok for &Get, but &Let.
'                          ! Because the prp is at at :ObjWiPrps level, not :Properties level.
On Error Resume Next
VzOPrps = PrpszO(ObjWiPrps)(P).Value
End Function

Private Sub Y_TblPrp()
Dim D As Database: Set D = TmpDb
DrpT D, "Tmp"
Rq D, "Create Table Tmp (F1 Text)"
TblPrp(D, "Tmp", "XX") = "AFdf"
Debug.Assert TblPrp(D, "Tmp", "XX") = "AFdf"
End Sub

Private Sub Z()
MDao_Z_Prp_Fld:
End Sub

Private Sub Z_FldPrp()
Dim P$, Db As Database, T, F$, V
GoSub T0
Exit Sub
T0:
    Set Db = TmpDb
    Rq Db, "Create Table Tmp (AA Text)"
    T = "Tmp"
    F = "AA"
    P = "Ele"
    V = "Ele1234"
    GoTo Tst
Tst:
    FldPrp(Db, T, F, P) = V
    Ass FldPrp(Db, T, F, P) = V
    Dim Fd As dao.Field: Set Fd = FdzTF(Db, T, F)
    Stop
    DmpDy PrpDyoFd(Fd)
    Return
End Sub

Private Sub Z_PrpDyoFd()
Dim Db As Database: Set Db = SampDbDutyDta
Dim Fd As dao.Field
Dim Rs As dao.Recordset
Set Rs = RszT(Db, "Permit")
Set Fd = Rs.Fields("Permit")
Debug.Print Fd.Value
DmpDy PrpDyoFd(Fd)
End Sub

Private Sub Z_PrpNy()
Dim Db As Database: Set Db = SampDbDutyDta
Dim Fd As dao.Field
Set Fd = FdzTF(Db, "Permit", "Permit")
D PrpNyzFd(Fd)
End Sub