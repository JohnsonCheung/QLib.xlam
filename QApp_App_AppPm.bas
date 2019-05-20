Attribute VB_Name = "QApp_App_AppPm"
Option Compare Text
Option Explicit
Private Const CMod$ = "MApp_Pm."
Private Const Asm$ = "QApp"

Function OupPth$(A As Database)
Dim P$: P = ValzPm(A, "OupPth")
EnsPthzAllSeg P
OupPth = P
End Function

Function PthzPm$(A As Database, PmNm$)
PthzPm = EnsPthSfx(ValzPm(A, PmNm & "Pth"))
End Function

Function Pjfnm$(A As Database, PmNm$)
Pjfnm = ValzPm(A, PmNm & "Fn")
End Function

Function FfnzPm(A As Database, PmNm$)
FfnzPm = PthzPm(A, PmNm) & Pjfnm(A, PmNm)
End Function

Property Get ValzPm$(A As Database, PmNm$)
Dim Q$: Q = FmtQQ("Select ? From Pm where CUsr='?'", PmNm, CUsr)
ValzPm = ValzQ(A, Q)
End Property

Property Let ValzPm(A As Database, PmNm$, V$)
With A.TableDefs("Pm").OpenRecordset
    .Edit
    .Fields(PmNm).Value = V
    .Update
End With
End Property
Sub BrwPm(A As Database)
BrwTbl A, "Pm"
End Sub

Private Sub ZZ()
End Sub

