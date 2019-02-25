Attribute VB_Name = "MDao_Att_Op_Dlt"
Option Explicit

Sub DltAttDb(A As Database, Att)
A.Execute FmtQQ("Delete * from Att where AttNm='?'", Att)
End Sub

Sub DltAtt(Att)
DltAttDb CDb, Att
End Sub


