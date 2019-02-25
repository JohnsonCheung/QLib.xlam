Attribute VB_Name = "MApp_Pm"
Option Explicit
Property Get PnmOupPthzDb$(Db As Database)
PnmOupPthzDb = PnmValz(Db, "OupPth")
End Property
Property Get PnmOupPth$()
PnmOupPth = PnmOupPthzDb(CDb)
End Property
Function PnmPthzDb$(Db As Database, A)
PnmPthzDb = PthEnsSfx(PnmValz(Db, A & "Pth"))
End Function

Function PnmFnz$(Db As Database, A)
PnmFnz = PnmValz(Db, A & "Fn")
End Function

Function PnmFfnz(Db As Database, A)
PnmFfnz = PnmPthzDb(Db, A) & PnmFnz(Db, A)
End Function
Function PnmFfn$(A)
PnmFfn = PnmFfnz(CDb, A)
End Function

Function PnmFn$(A)
PnmFn = PnmVal(A & "Fn")
End Function

Function PnmPth$(A)
PnmPth = PthEnsSfx(PnmVal(A & "Pth"))
End Function

Property Get PnmVal$(Pnm$)
PnmVal = PnmValz(CDb, Pnm)
End Property

Property Get PnmValz$(Db As Database, Pnm$)
PnmValz = ValzDbtf(Db, "Pm", Pnm)
End Property

Property Let PnmVal(Pnm$, V$)
Stop
'Should not use
With CDb.TableDefs("Pm").OpenRecordset
    .Edit
    .Fields(Pnm).Value = V
    .Update
End With
End Property

Sub BrwTblPm()
'Tbl_Opn "Pm"
End Sub

Private Sub ZZ()
Dim A$
Dim B As Variant
PnmFfn A
PnmFfn B
PnmFn B
PnmPth B
End Sub

Private Sub Z()
End Sub
