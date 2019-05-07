Attribute VB_Name = "QDao_Dic"
Option Explicit
Private Const CMod$ = "MDao_Dic."
Private Const Asm$ = "QDao"


Function AyDaoTy(A As DAO.DataTypeEnum)
Dim O
Select Case A
Case DAO.DataTypeEnum.dbBigInt: O = EmpLngAy
End Select
End Function
Function AyDic_RsKF(A As DAO.Recordset, DicKeyFld, AyFld) As Dictionary _
'Return a dictionary of Ay using KeyFld and AyFld.  The Val-of-returned-Dic is Ay using the AyFld.Type to create
Dim O As New Dictionary
Dim K, V
Dim Emp
Dim Ay
    Emp = AyDaoTy(A.Fields(AyFld).Type)
    Ay = Emp
With A
    While Not .EOF
        K = .Fields(DicKeyFld).Value
        V = .Fields(AyFld).Value
        If O.Exists(K) Then
            If True Then
                Ay = O(K)
                PushI Ay, V
                O(K) = Ay
            Else
                PushI O(K), V '<-- It does not work
            End If
        Else
            Ay = Emp
            PushI Ay, V
            O.Add K, Ay
        End If
        .MoveNext
    Wend
End With
Set AyDic_RsKF = O
End Function


Function JnStrDicTwoFldRs(A As DAO.Recordset, Optional Sep$ = " ") As Dictionary
Set JnStrDicTwoFldRs = JnStrDicRsKeyJn(A, 0, 1, Sep)
End Function

Function JnStrDicRsKeyJn(A As DAO.Recordset, KeyFld, JnStrFld, Optional Sep$ = " ") As Dictionary
Dim O As New Dictionary
Dim K, V$
While Not A.EOF
    K = A.Fields(KeyFld).Value
    V = Nz(A.Fields(JnStrFld).Value, "")
    If O.Exists(K) Then
        O(K) = O(K) & Sep & V
    Else
        O.Add K, CStr(Nz(V))
    End If
    A.MoveNext
Wend
Set JnStrDicRsKeyJn = O
End Function

Function CntDiczRs(A As DAO.Recordset, Optional Fld = 0) As Dictionary
Set CntDiczRs = CntDic(AvRsCol(A))
End Function

