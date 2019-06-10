Attribute VB_Name = "QIde_Mth"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth."
Private Const Asm$ = "QIde"
Function AA() As String()
Erase XX
X "skldfjls dklskdjfl kdfj klsdfjfjlsdkf"
X "sdflksjdf"
X ""
X "skldjflskdjfsdf"
AA = XX
End Function

Property Get MthKeyFny() As String()
MthKeyFny = SyzSS("Pjn Mdn Priority Nm Ty Mdy")
End Property

Function RplMth(M As CodeModule, Mthn, NewL$) As Boolean
'Ret True if Rplaced
Dim Lno&: Lno = MthLnozMM(M, Mthn)
If Not HasMthzM(M, Mthn) Then
    RplMth = True
    M.AddFromString NewL '<===
    Exit Function
End If
Dim OldL$: OldL = MthLineszM(M, Mthn)
If OldL = NewL Then Exit Function
RplMth = True
RmvMth M, Mthn '<==
M.InsertLines Lno, NewL '<==
End Function

Private Sub ZZ()
MIde__Mth:
End Sub

Sub EnsLines(Md As CodeModule, Mthn, MthLines$)
Dim OldMthLines$: OldMthLines = MthLineszM(Md, Mthn)
If OldMthLines = MthLines Then
    Debug.Print FmtQQ("EnsMd: Mth(?) in Md(?) is same", Mthn, Mdn(Md))
End If
RmvMthzMN Md, Mthn
ApdLines Md, MthLines
Debug.Print FmtQQ("EnsMd: Mth(?) in Md(?) is replaced <=========", Mthn, Mdn(Md))
End Sub

