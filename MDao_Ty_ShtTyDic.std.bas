Attribute VB_Name = "MDao_Ty_ShtTyDic"
Option Explicit
Function ShtTyDic(Ffn, TblNm) As Dictionary
Select Case True
Case IsFb(Ffn): Set ShtTyDic = ShtTyDiczFbt(Ffn, TblNm)
Case IsFx(Ffn): Set ShtTyDic = ShtTyDiczFxw(Ffn, TblNm)
Case Else: Thw CSub, "Ffn should be Fx or Fb", "Ffn TblNm", Ffn, TblNm
End Select
End Function

Private Function ShtTyDiczFbt(Fb, T) As Dictionary
Dim F As Dao.Field
Set ShtTyDiczFbt = New Dictionary
For Each F In Db(Fb).TableDefs(T).Fields
    ShtTyDiczFbt.Add F.Name, ShtTyzDao(F.Type)
Next
End Function

Function ShtTyDiczFxw(Fx, W) As Dictionary
Dim C As Column, Cat As Catalog, I
Set Cat = CatzFx(Fx)
For Each I In Cat.Tables(CatT(W)).Columns
    ShtTyDiczFxw.Add C.Name, ShtTyzAdo(C.Type)
Next
End Function


