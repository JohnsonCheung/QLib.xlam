Attribute VB_Name = "MXls_Rfh"
Option Explicit
Private Sub ClsWc(A As WorkbookConnection)
If IsNothing(A.OLEDBConnection) Then Exit Sub
CvCn(A.ODBCConnection.Connection).Close
End Sub

Private Sub ClsWczWb(Wb As Workbook)
Dim Wc As WorkbookConnection
For Each Wc In Wb.Connections
    ClsWc Wc
Next
End Sub

Private Sub SetWczFb(A As WorkbookConnection, ToUseFb)
If IsNothing(A.OLEDBConnection) Then Exit Sub
Dim Cn$
#Const A = 2
#If A = 1 Then
    Dim S$
    S = A.OLEDBConnection.Connection
    Cn = RplBet(S, ToUseFb, "Data Source=", ";")
#End If
#If A = 2 Then
    Cn = CnStrzFbAdoOle(ToUseFb)
#End If
A.OLEDBConnection.Connection = Cn
End Sub

Private Sub RfhWc(A As WorkbookConnection, ToUseFb)
If IsNothing(A.OLEDBConnection) Then Exit Sub
SetWczFb A, ToUseFb
A.OLEDBConnection.BackgroundQuery = False
A.OLEDBConnection.Refresh
End Sub

Private Sub RfhPc(A As PivotCache)
A.MissingItemsLimit = xlMissingItemsNone
A.Refresh
End Sub

Sub RfhFx(Fx$, Fb$)
RfhWb(WbzFx(Fx$), Fb).Close SaveChanges:=True
End Sub

Private Sub RfhWs(A As Worksheet)
Dim Q As QueryTable: For Each Q In A.QueryTables: Q.BackgroundQuery = False: Q.Refresh: Next
Dim P As PivotTable: For Each P In A.PivotTables: P.Update: Next
Dim L As ListObject: For Each L In A.ListObjects: L.Refresh: Next
End Sub

Function RfhWb(Wb As Workbook, Fb) As Workbook
RplLozFb Wb, Fb
Dim C As WorkbookConnection
Dim P As PivotCache, W As Worksheet
For Each C In Wb.Connections: RfhWc C, Fb:                                          Next
For Each P In Wb.PivotCaches: P.MissingItemsLimit = xlMissingItemsNone: P.Refresh:  Next
For Each W In Wb.Sheets:      RfhWs W:                                              Next
FmtLozStdWb Wb
ClsWczWb Wb
DltWc Wb
Set RfhWb = Wb
End Function

Private Sub RplLozFb(Wb As Workbook, Fb)
Dim I, Lo As ListObject, D As Database
Set D = Db(Fb$)
For Each I In OupLoAy(Wb)
    Set Lo = I
    RplLozT Lo, D, "@" & Mid(Lo.Name, 3)
Next
D.Close
Set D = Nothing
End Sub

Private Function RplLozT(A As ListObject, Db As Database, T$) As ListObject
Dim Fny1$(): Fny1 = Fny(Db, T)
Dim Fny2$(): Fny2 = FnyzLo(A)
If Not IsSamAy(Fny1, Fny2) Then
    Thw CSub, "LoFny and TblFny are not same", "LoFny TblNm TblFny Db", Fny2, T, Fny1, DbNm(A)
End If
Dim Sq()
    Dim R As DAO.Recordset
    Set R = Rs(A, SqlSel_FF_Fm(Fny2, T))
    Sq = SqAddSngQuote(SqzRs(R))
MinxLo A
RgzSq Sq, A.DataBodyRange
Set RplLozT = A
End Function

Private Function OupLoAy(A As Workbook) As ListObject()
OupLoAy = OywNmPfx(LoAy(A), "T_")
End Function


