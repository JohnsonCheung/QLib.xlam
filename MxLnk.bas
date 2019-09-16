Attribute VB_Name = "MxLnk"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxLnk."

Sub LnkTblzDrs(D As Database, DoTSCn As Drs)
LnkTblzDy D, DoTSCn.Dy
End Sub

Sub LnkTblzDy(D As Database, DyoTSCn())
Dim Dr, T$, S$, CN$
For Each Dr In Itr(DyoTSCn)
    T = Dr(0)
    S = Dr(1)
    CN = Dr(2)
    LnkTbl D, T, S, CN
Next
End Sub

Sub LnkTbl(D As Database, T, S$, CN$)
On Error GoTo X
DrpT D, T
D.TableDefs.Append TdzCnStr(T, S, CN)
Exit Sub
X:
    Dim Er$: Er = Err.Description
    Thw CSub, "Error in linking table", "Er Db T SrcTbl Cn", Er, D.Name, T, S, CN
End Sub

Function ErzLnkFxw(D As Database, T, Fx, Optional Wsn = "Sheet1") As String()
On Error GoTo X
LnkFxw D, T, Fx, Wsn
Exit Function
X: ErzLnkFxw = _
    LyzMsgNap("Error in linking Xls file", "Er LnkFx LnkWs ToDb AsTbl", Err.Description, Fx, Wsn, D.Name, T)
End Function

Sub LnkFxw(D As Database, T, Fx, Optional Wsn = "Sheet1")
LnkTbl D, T, Wsn & "$", DaoCnStrzFx(Fx)
End Sub

Sub LnkFb(D As Database, T, Fb, Optional Fbt$)
LnkTbl D, T, DftStr(Fbt, T), DaoCnStrzFb(Fb)
End Sub

Private Function TdzCnStr(T, Src$, CN$) As dao.TableDef
Set TdzCnStr = New dao.TableDef
With TdzCnStr
    .Connect = CN
    .Name = T
    .SourceTableName = Src
End With
End Function
Function CnStrAy(D As Database) As String()
Dim T: For Each T In Tni(D)
    PushNB CnStrAy, CnStrzT(D, T)
Next
End Function