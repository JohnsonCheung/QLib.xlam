Attribute VB_Name = "QDao_Ty_ShtTyDic"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Ty_ShtTyDic."
Private Const Asm$ = "QDao"

Function ShtTyDic(FxOrFb$, TblNm$) As Dictionary
Select Case True
Case IsFb(FxOrFb): Set ShtTyDic = ShtTyDiczFbt(FxOrFb, TblNm$)
Case IsFx(FxOrFb): Set ShtTyDic = ShtTyDiczFxw(FxOrFb, TblNm$)
Case Else: Thw CSub, "FxOrFb should be Fx or Fb", "FxOrFb TblNm", FxOrFb, TblNm
End Select
End Function

Private Function ShtTyDiczFbt(Fb, T) As Dictionary
Dim F As DAO.Field
Set ShtTyDiczFbt = New Dictionary
For Each F In Db(Fb).TableDefs(T).Fields
    ShtTyDiczFbt.Add F.Name, ShtTyzDao(F.Type)
Next
End Function

Private Function ShtTyDiczFxw(Fx, W$) As Dictionary
Dim C As Column, Cat As Catalog, I
Set Cat = CatzFx(Fx)
For Each I In Cat.Tables(CatTn(W)).Columns
    ShtTyDiczFxw.Add C.Name, ShtTyzAdo(C.Type)
Next
End Function


