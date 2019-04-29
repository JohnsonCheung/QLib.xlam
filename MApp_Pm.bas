Attribute VB_Name = "MApp_Pm"
Option Explicit

Function PmOupPth$(A As Database)
PmOupPth = EnsPthzAllSeg(PmVal(Db, "OupPth"))
End Function

Function PmPth$(A As Database, PmNm$)
PmPth = EnsPthSfx(PmVal(Db, PmNm & "Pth"))
End Function

Function PmFn$(A As Database, PmNm$)
PmFn = PmVal(Db, PmNm & "Fn")
End Function

Function PmFfn(A As Database, PmNm$)
PmFfn = PmPth(Db, PmNm) & PmFn(Db, PmNm)
End Function

Property Get PmVal$(A As Database, PmNm$)
PmVal = ValOfTF(Db, "Pm", PmNm)
End Property

Property Let PmVal(A As Database, PmNm$, V$)
With A.TableDefs("Pm").OpenRecordset
    .Edit
    .Fields(PmNm).Value = V
    .Update
End With
End Property
Sub BrwPm(A As Database)
BrwTbl Db, "Pm"
End Sub

Private Sub ZZ()
End Sub

Private Sub Z()
End Sub
