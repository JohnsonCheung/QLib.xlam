Attribute VB_Name = "QDao_B_Prp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Prp."
Private Const Asm$ = "QDao"

Function FldDes$(D As Database, T, F$)
FldDes = FldPrp(D, T, F, C_Des)
End Function

Sub SetFldDes(D As Database, T, F$, Des$)
FldPrp(D, T, F, C_Des) = Des
End Sub

Private Sub Z_PrpNy()
Dim Db As Database: Set Db = SampDboDutyDta
Dim Fd As Dao.Field
Set Fd = FdzTF(Db, "Permit", "Permit")
D PrpNyzFd(Fd)
End Sub

Function PrpNyzFd(A As Dao.Field) As String()
PrpNyzFd = Itn(A.Properties)
End Function

Private Sub Z_PrpDyoFd()
Dim Db As Database: Set Db = SampDboDutyDta
Dim Fd As Dao.Field
Dim Rs As Dao.Recordset
Set Rs = RszT(Db, "Permit")
Set Fd = Rs.Fields("Permit")
Debug.Print Fd.Value
DmpDy PrpDyoFd(Fd)
End Sub

Function PrpDyoFd(A As Dao.Field) As Variant()
Dim PrpV, I, P$, V
For Each I In Itn(A.Properties)
    P = I
    V = VzOPrps(A, P)
    PushI PrpDyoFd, Array(P, V, TypeName(V))
Next
End Function

Sub SetVzOPrps(ObjWiPrps, P$, V)
Dim Prps As Dao.Properties: Set Prps = PrpszO(ObjWiPrps)
If HasItn(Prps, P) Then
    Prps(P).Value = V
Else
    Prps.Append ObjWiPrps.CreateProperty(P, DaoTyzV(V), V) ' will break if V=""
End If
End Sub

Function VzOPrps(ObjWiPrps, P$)
'Ret : #Val-fm-ObjWithPrps ! Notes: Just passing @ObjWiPrps.Properties is Ok for &Get, but &Let.
'                          ! Because the prp is at at :ObjWiPrps level, not :Properties level.
On Error Resume Next
VzOPrps = PrpszO(ObjWiPrps)(P).Value
End Function

Function PrpszO(ObjWiPrps) As Dao.Properties
On Error GoTo X
Set PrpszO = ObjWiPrps.Properties
Exit Function
X:
    Dim E$: E = Err.Description
    Thw CSub, "Obj does not have prp-[Properties]", "Obj-TyNm Er", TypeName(ObjWiPrps), E
End Function

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
    Dim Fd As Dao.Field: Set Fd = FdzTF(Db, T, F)
    Stop
    DmpDy PrpDyoFd(Fd)
    Return
End Sub

Function FldDeszTd$(A As Dao.Field)
FldDeszTd = VzOPrps(A.Properties, C_Des)
End Function

Sub SetFldDeszTd(A As Dao.Field, Des$)

End Sub

Private Sub Z()
MDao_Z_Prp_Fld:
End Sub

Private Sub Y_TblPrp()
Dim D As Database: Set D = TmpDb
DrpT D, "Tmp"
Rq D, "Create Table Tmp (F1 Text)"
TblPrp(D, "Tmp", "XX") = "AFdf"
Debug.Assert TblPrp(D, "Tmp", "XX") = "AFdf"
End Sub

Function HasDbtPrp(D As Database, T, P) As Boolean
HasDbtPrp = HasItn(D.TableDefs(T).Properties, P)
End Function

Function TblPrp(D As Database, T, P)
If Not HasDbtPrp(D, T, P) Then Exit Function
TblPrp = D.TableDefs(T).Properties(P).Value
End Function

Sub SetTblPrp(D As Database, T, P, V)
Dim Td As Dao.TableDef: Set Td = D.TableDefs(T)
SetDaoPrp Td, Td.Properties, P, V
End Sub

Sub SetDaoPrp(WiDaoPrps As Object, Prps As Dao.Properties, P, V)
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

Sub SetFldPrp(D As Database, T, F$, P$, V)
Dim Fd As Dao.Field: Set Fd = D.TableDefs(T).Fields(F)
SetDaoPrp Fd, Fd.Properties, P, V
End Sub

Function FldPrp(D As Database, T, F$, P$)
If Not HasFldPrp(D, T, F, P) Then Exit Function
FldPrp = D.TableDefs(T).Fields(F).Properties(P).Value
End Function

Function HasFldPrp(D As Database, T, F$, P$) As Boolean
HasFldPrp = HasItn(D.TableDefs(T).Fields(F).Properties, P)
End Function

