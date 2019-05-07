Attribute VB_Name = "QDao_Sq"
Const Asm$ = "Dao"
Const Ns$ = "Dao.Sq"
Private Const CMod$ = "BSq_Dao."
Sub SetSqrzRs(OSq(), R&, A As DAO.Recordset, Optional NoTxtSngQ As Boolean)
SetSqzDrv OSq, R, DrzRs(A), NoTxtSngQ
End Sub
Sub SetSqzDrv(OSq(), R&, Drv, Optional NoTxtSngQ As Boolean)
Dim J&
If NoTxtSngQ Then
    For J = 0 To UB(Drv)
        If IsStr(Drv(J)) Then
            OSq(R, J + 1) = QuoteSng(CStr(Drv(J)))
        Else
            OSq(R, J + 1) = Drv(J)
        End If
    Next
Else
    For J = 0 To UB(Drv)
        OSq(R, J + 1) = Drv(J)
    Next
End If
End Sub

