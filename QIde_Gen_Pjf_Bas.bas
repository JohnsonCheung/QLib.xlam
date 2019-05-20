Attribute VB_Name = "QIde_Gen_Pjf_Bas"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Gen_Pjf_Bas."
Private Const Asm$ = "QIde"
Sub LoadBas(P As VBProject, Srcp$)
Dim BasItm
For Each BasItm In Itr(BasFfny(Srcp))
    P.VBComponents.Import BasItm
Next
End Sub

Private Function BasFfny(Srcp$) As String()
Dim Ffn$, I
For Each I In Itr(Ffny(Srcp))
    Ffn = I
    If IsBasFfn(Ffn) Then
        PushI BasFfny, Ffn
    End If
Next
End Function
Private Function IsBasFfn(Ffn) As Boolean
Select Case True
Case HasSfx(Ffn, ".bas"): IsBasFfn = True
End Select
End Function

