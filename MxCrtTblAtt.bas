Attribute VB_Name = "MxCrtTblAtt"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxCrtTblAtt."
Sub CrtTblAtt(D As Database)
Dim PFldCsv$: PFldCsv = "Attk Text(255), Att Attachment, FilSi Long,FilTim Date" ' #Sql-Fld-Phrase.  The fld spec of create table sql inside the bkt.
CrtTblzPFld D, "Att", PFldCsv
DbzReOpn(D).Execute "Create Index PrimaryKey on Att (Attk) with Primary"
End Sub

Sub EnsTblAtt(D As Database)
If HasTbl(D, "Att") Then
    Dim FF$: FF = FFzT(D, "Att")
    If FF <> "Attk Att FilSi FilTim" Then Thw CSub, "Db has :Tbl:Att, but its FF is not [Attk Att FilSi FilTim", "Dbn Tbl-Att-FF", D.Name, FF
End If
CrtTblAtt D
End Sub

Sub Z_EnsTblAtt()
Dim D As Database: Set D = TmpDb
EnsTblAtt D
BrwDb D
End Sub

