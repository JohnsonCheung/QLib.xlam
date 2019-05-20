Attribute VB_Name = "QDao_Att_Op_Dlt"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Att_Op_Dlt."
Private Const Asm$ = "QDao"

Sub DltAtt(A As Database, Att$)
A.Execute FmtQQ("Delete * from Att where AttNm='?'", Att)
End Sub


