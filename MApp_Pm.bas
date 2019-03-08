Attribute VB_Name = "MApp_Pm"
Option Explicit
Function PnmOupPth$(A As Database)
PnmOupPth = PthEnsAll(PnmVal(A, "OupPth"))
End Function

Function PnmPth$(Db As Database, Pnm)
PnmPth = PthEnsSfx(PnmVal(Db, Pnm & "Pth"))
End Function

Function PnmFn$(Db As Database, Pnm)
PnmFn = PnmVal(Db, Pnm & "Fn")
End Function

Function PnmFfn(Db As Database, Pnm)
PnmFfn = PnmPth(Db, Pnm) & PnmFn(Db, Pnm)
End Function

Property Get PnmVal$(Db As Database, Pnm$)
PnmVal = ValzTF(Db, "Pm", Pnm)
End Property

Property Let PnmVal(Db As Database, Pnm$, V$)
Stop
'Should not use
With Db.TableDefs("Pm").OpenRecordset
    .Edit
    .Fields(Pnm).Value = V
    .Update
End With
End Property

Sub BrwTblPm(Apn$)
BrwTbl AppDb(Apn), "Pm"
End Sub

Private Sub ZZ()
End Sub

Private Sub Z()
End Sub
