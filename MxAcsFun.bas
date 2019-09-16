Attribute VB_Name = "MxAcsFun"
Option Compare Text
Option Explicit
Const CLib$ = "QAcs."
Const CMod$ = CLib & "MxAcsFun."

Function Acs() As Access.Application
Set Acs = Access.Application
End Function

Function AcszDb(D As Database) As Access.Application
Static A As New Access.Application
OpnFb A, D.Name
Set AcszDb = A
A.Visible = True
End Function

Sub BrwFb(Fb)
Static Acs As New Access.Application
OpnFb Acs, Fb
Acs.Visible = True
End Sub

Sub BrwT(D As Database, T)
AcszDb(D).DoCmd.OpenTable T
End Sub

Sub BrwTT(D As Database, TT$)
Dim T
For Each T In ItrzTT(TT)
    BrwT D, T
Next
End Sub

Sub ClsAllTbl(D As Database)
Dim A As Access.Application: Set A = AcszDb(D)
Dim T: For Each T In A.CodeData.AllTables
    ClsTzA A, T
Next
End Sub

Sub ClsDbzAcs(A As Access.Application)
On Error Resume Next
A.CloseCurrentDatabase
End Sub

Sub ClsT(D As Database, T)
AcszDb(D).DoCmd.Close acTable, T
End Sub

Sub ClsTT(D As Database, TT$)
Dim A As Access.Application: Set A = AcszDb(D)
Dim T: For Each T In TermAy(TT)
    ClsTzA A, T
Next
End Sub

Sub ClsTzA(A As Access.Application, T)
A.DoCmd.Close acTable, T
End Sub

Sub CpyAcsObj(A As Access.Application, ToFb)
ThwIf_FfnExist ToFb, CSub, "ToFb"
CrtFb ToFb
CpyTzA A, ToFb
CpyFrm A, ToFb
CpyMdzA A, ToFb
CpyRfzA A, ToFb
End Sub

Sub CpyFrm(A As Access.Application, Fb)
Dim I As AccessObject: For Each I In A.CodeProject.AllForms
    A.DoCmd.CopyObject Fb, , acForm, I.Name
Next
End Sub

Sub CpyMdzA(A As Access.Application, ToFb)
Dim I As AccessObject
For Each I In A.CodeProject.AllModules
    A.DoCmd.CopyObject ToFb, , acModule, I.Name
Next
End Sub

Sub CpyRfzA(A As Access.Application, ToFb)

End Sub

Sub CpyTzA(A As Access.Application, ToFb)
Dim I As AccessObject: For Each I In A.CodeData.AllTables
    A.DoCmd.CopyObject ToFb, , acTable, I.Name
Next
End Sub

Function CvAcs(A) As Access.Application
Set CvAcs = A
End Function

Function DbnzAcs$(A As Access.Application)
On Error Resume Next
DbnzAcs = A.CurrentDb.Name
End Function

Function DftAcs(A As Access.Application) As Access.Application
'Ret :@A if Not Nothing or :NewAcs
If IsNothing(A) Then
    Set DftAcs = NewAcs
Else
    Set DftAcs = A
End If
End Function

Function FbzAcs$(A As Access.Application)
'Ret :Dbn openned in @A or *Blnk
On Error Resume Next
FbzAcs = A.CurrentDb.Name
End Function

Function NewAcs(Optional Shw As Boolean) As Access.Application
Dim O As Access.Application: Set O = CreateObject("Access.Application")
If Shw Then O.Visible = True
Set NewAcs = O
End Function

Sub OpnFb(A As Access.Application, Fb)
'Do : Opn @Fb in @A @@
If DbnzAcs(A) = Fb Then Exit Sub
ClsDbzAcs A
A.OpenCurrentDatabase Fb
End Sub

Function PjzAcs(A As Access.Application) As VBProject
Set PjzAcs = A.Vbe.ActiveVBProject
End Function

Function PjzFba(Fba, A As Access.Application) As VBProject
OpnFb A, Fba
Set PjzFba = PjzAcs(A)
End Function

Sub QuitAcs(A As Access.Application)
If IsNothing(A) Then Exit Sub
On Error Resume Next
Stamp "QuitAcs: Begin"
Stamp "QuitAcs: Cls":         A.CloseCurrentDatabase
Stamp "QuitAcs: Quit":        A.Quit
Stamp "QuitAcs: Set Nothing": Set A = Nothing
Stamp "QuitAcs: End"
End Sub

Sub SavRec()
DoCmd.RunCommand acCmdSaveRecord
End Sub

Function ShwAcs(A As Access.Application) As Access.Application
If Not A.Visible Then A.Visible = True
Set ShwAcs = A
End Function