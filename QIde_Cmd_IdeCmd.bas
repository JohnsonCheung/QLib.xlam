Attribute VB_Name = "QIde_Cmd_IdeCmd"
Option Explicit
Option Compare Text
Sub Cmd_AlignMth()
Dim IsAlignable As Boolean, Lno&, NLin%, OldL$, NewL$, Md As CodeModule, MthLy$(), OldMthLy$(), N$
     Set Md = CMd
        Lno = MthLno(Md, CLno)
          N = Mthn(Md.Lines(Lno, 1))
              If N = "Cmd_AlignMth" Then MsgBox "Cannot align Mth: Cmd_AlignMth", vbCritical: Exit Sub
       NLin = NMthLin(Md, Lno)
       OldL = Md.Lines(Lno, NLin)
   OldMthLy = SplitCrLf(OldL)
IsAlignable = IsAlignableMth(OldMthLy)
              If Not IsAlignable Then MsgBox "Mth: " & N & vbCrLf & "is not Alignable", vbCritical: Exit Sub
       NewL = AlignMth(OldMthLy)
       Stop
              'RplLines Md, Lno, NLin, OldL, NewL
End Sub

Private Sub Z_AlignMthL()
Dim M$()
GoSub Z
Exit Sub
Z:
    Erase XX
    X "Sub Cmd_AlignMth()"
    X "Dim IsAlignable As Boolean, Lno&, NLin%, OldL$, NewL$, Md As CodeModule, MthLy$()"
    X "Set Md = CMd"
    X "Lno = MthLno(Md, CLno)"
    X "Stop"
    X "NLin = NMthLin(Md, Lno)"
    X "OldL = Md.Lines(Lno, NLin)"
    X "Nm = Mthn(Md.Lines(Lno, 1))"
    X "IsAlignable = IsAlignableMth(OldL)"
    X "If Not IsAlignable Then MsgBox ""Mth: "" & Mthn & vbCrLf & ""Alignable"", vbCritical: Exit Sub"
    X "NewL = JnCrLf(AlignMthL(MthL))"
    X "RplLines Md, Lno, NLin, OldL, NewL"
    X "End Sub"
    M = XX
    Erase XX

    D AlignMth(M)
    Return
End Sub
Private Function IsAlignableMth(MthLy$()) As Boolean
Dim L
For Each L In MthCxtLy(MthLy)
    If Not IsLinAlignableSrc(L) Then Exit Function
Next
IsAlignableMth = True
End Function
Private Function AlignMth$(MthLy$())
Dim ONoAlign$(), OLHS$(), ORHS$(), O$(), U%, J%, Fm%, L$
U = UB(MthLy)
ReDim ONoAlign(U), OLHS(U), ORHS(U), O(U)
O(U) = MthLy(U)
Fm = NxtIxzSrc(MthLy)
For J = 0 To Fm - 1
    O(J) = MthLy(J)
Next
For J = Fm To U - 1
    L = MthLy(J)
    If ShouldAlign(L) Then
        With Brk2(L, " = ", NoTrim:=True)
            OLHS(J) = ApdIf(.S1, " = ")
            ORHS(J) = .S2
        End With
    Else
        ONoAlign(J) = L
    End If
Next
OLHS = AlignRzAy(OLHS)
For J = Fm To U - 1
    Select Case True
    Case ONoAlign(J) <> "": O(J) = ONoAlign(J)
    Case Else:              O(J) = OLHS(J) & ORHS(J)
    End Select
Next
AlignMth = JnCrLf(O)
End Function
Private Function ShouldAlign(L$) As Boolean
If T1(L) = "Dim" Then Exit Function
If FstChr(LTrim(L)) = "'" Then Exit Function
ShouldAlign = True
End Function
Private Function IsLinAlignableSrc(L) As Boolean
Dim T1$
T1 = T1zS(L)
     Select Case T1
     Case "Dim", "Set", "If", "On"
        IsLinAlignableSrc = True
        Exit Function
     Case "Select", "Case", "End", "Else", "For", "While", "Do"
        Exit Function
     End Select
     If LasChr(T1) = ":" Then Exit Function
     IsLinAlignableSrc = True
End Function
