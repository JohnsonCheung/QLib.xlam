Attribute VB_Name = "MVb_Tst"
Option Explicit
Public Act, Ept, Dbg As Boolean, Trc As Boolean
Sub StopNE()
If Not IsEq(Act, Ept) Then Stop
End Sub
Sub C()
ThwNE Act, Ept
End Sub

Sub BrwTstPth(Fun$, Cas$)
BrwPth TstPth(Fun, Cas)
End Sub

Private Function TstPth$(Fun$, Cas$)
TstPth = TstHom & JnPthSepAp(Replace(Fun, ".", PthSep), Cas) & PthSep
End Function

Property Get TstHom$()
Static X$
Dim P$
P = Pth(Application.Vbe.ActiveVBProject.Filename)
If X = "" Then X = PthEns(P & "TstRes\")
TstHom = X
End Property

Sub BrwTstHom()
BrwPth TstHom
End Sub

Sub ShwTstOk(Fun$, Cas$)
Debug.Print "Tst OK | "; Fun; " | Case "; Cas
End Sub

Function TstTxt$(Fun$, Cas$, Itm$, Optional IsEdt As Boolean)
If IsEdt Then
    EdtTstTxt Fun, Cas, Itm
    Exit Function
End If
TstTxt = FtLines(TstFt(Fun, Cas, Itm))
End Function

Sub EdtTstTxt(Fun$, Cas$, Itm$)
BrwFt TstFt(Fun, Cas, Itm)
End Sub

Private Function TstFt$(Fun$, Cas$, Itm$)
TstFt = PthEnsAll(TstPth(Fun, Cas) & Itm & ".txt")
End Function
