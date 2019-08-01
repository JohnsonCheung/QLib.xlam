Attribute VB_Name = "QIde_F_Res"

Function ResLyzMthn(M As CodeModule, Mthn$) As String()
Dim Z$
    Z = MthLzM(M, Mthn)
    If Si(Z) = 0 Then
        Thw CSub, "Mthn not found", "Mthn Md", Mthn, Mdn(M)
    End If
    Z = AeFstEle(Z)
    Z = AeLasEle(Z)
ResLyzMthn = RmvFstChrzAy(Z)
End Function

Function ReszMthn$(M As CodeModule, Mthn$)
ReszMthn = JnCrLf(ResLyzMthn(M, Mthn))
End Function

