Attribute VB_Name = "MDao_Att_Op_Dlt"
Option Explicit

Sub DltAtt(A As Database, Att$)
A.Execute FmtQQ("Delete * from Att where AttNm='?'", Att)
End Sub


