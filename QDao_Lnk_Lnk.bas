Attribute VB_Name = "QDao_Lnk_Lnk"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Lnk."
Private Const Asm$ = "QDao"

Sub LnkTbl(A As Database, T, S$, Cn$)
On Error GoTo X
DrpT A, T
A.TableDefs.Append TdzzTSC(T, S, Cn)
Exit Sub
X:
    Dim Er$: Er = Err.Description
    Thw CSub, "Error in linking table", "Er Db T SrcTbl Cn", Er, Dbn(A), T, S, Cn
End Sub

Sub LnkFx(A As Database, T, Fx, Optional Wsn = "Sheet1")
LnkTbl A, T, Wsn & "$", CnStrzFxDao(Fx)
End Sub

Sub LnkFb(A As Database, T, Fb, Optional Fbt$)
LnkTbl A, T, DftStr(Fbt, T), CnStrzFbDao(Fb)
End Sub

Private Function TdzzTSC(T, Src$, Cn$) As DAO.TableDef
Set TdzzTSC = New DAO.TableDef
With TdzzTSC
    .Connect = Cn
    .Name = T
    .SourceTableName = Src
End With
End Function
Function CnStrAy(D As Database) As String()
Dim T
For Each T In Tbli(D)
    PushNonBlank CnStrAy, CnStrzT(D, T)
Next
End Function
Function LnkgTny(A As Database) As String()
Dim T As TableDef
For Each T In A.TableDefs
    If T.Connect <> "" Then
        PushI LnkgTny, T.Name
    End If
Next
End Function




