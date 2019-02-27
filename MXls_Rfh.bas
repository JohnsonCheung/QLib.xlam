Attribute VB_Name = "MXls_Rfh"
Option Explicit
Sub ClsWcvWc(A As WorkbookConnection)
If IsNothing(A.OLEDBConnection) Then Exit Sub
CvCn(A.ODBCConnection.Connection).Close
End Sub

Sub ClsWc(Wb As Workbook)
Dim Wc As WorkbookConnection
For Each Wc In Wb.Connections
    ClsWcvWc Wc
Next
End Sub

Sub SetWcFb(A As WorkbookConnection, ToUseFb$)
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
Sub RfhWc(A As WorkbookConnection, ToUseFb$)
If IsNothing(A.OLEDBConnection) Then Exit Sub
SetWcFb A, ToUseFb
A.OLEDBConnection.BackgroundQuery = False
A.OLEDBConnection.Refresh
End Sub

Sub RfhPc(A As PivotCache)
A.MissingItemsLimit = xlMissingItemsNone
A.Refresh
End Sub

Sub RfhFx(Fx, Fb$)
WbRfh(WbzFx(Fx), Fb).Close SaveChanges:=True
End Sub

Sub RfhWs(A As Worksheet)
DoItrFun A.QueryTables, "RfhQt"
DoItrFun A.PivotTables, "RfhPt"
DoItrFun A.ListObjects, "RfhLo"
End Sub

Sub RfhLo(A As ListObject)
A.Refresh
End Sub

Sub RfhQt(A As Excel.QueryTable)
A.BackgroundQuery = False
A.Refresh
End Sub

Function WbRfh(Wb As Workbook, Fb$) As Workbook
RfhWb Wb, Fb
Set WbRfh = Wb
End Function

Function RfhWb(Wb As Workbook, Fb) As Workbook
RplLoCn Wb, Fb
DoItrFunXP Wb.Connections, "RfhWc", Fb
DoItrFun Wb.PivotCaches, "RfhPc"
DoItrFun Wb.Sheets, "RfhWs"
'FmtLozWb Wb
Stop
Dim Wc As WorkbookConnection
For Each Wc In Wb.Connections
    DltWc Wb
Next
ClsWc Wb
Set RfhWb = Wb
End Function

Sub RfhPt(A As Excel.PivotTable)
A.Update
End Sub
