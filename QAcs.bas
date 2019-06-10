Attribute VB_Name = "QAcs"
Option Compare Text
Option Explicit
Private Const CMod$ = "MAcs."
Private Const Asm$ = "Q"

Sub BrwTbl(D As Database, T)
CAcs(D).DoCmd.OpenTable T
End Sub

Sub BrwTT(D As Database, TT$)
Dim T
For Each T In ItrzTT(TT)
    BrwTbl D, T
Next
End Sub

Function CAcs(D As Database) As Access.Application
Static A As New Access.Application
OpnFb A, D.Name
Set CAcs = A
A.Visible = True
End Function

Sub SavRec()
DoCmd.RunCommand acCmdSaveRecord
End Sub

Function FbzAcs$(A As Access.Application)
On Error Resume Next
FbzAcs = A.CurrentDb.Name
End Function

Sub ClsDbzAcs(A As Access.Application)
On Error Resume Next
A.CloseCurrentDatabase
End Sub

Sub BrwFb(FB)
Static Acs As New Access.Application
OpnFb Acs, FB
Acs.Visible = True
End Sub

Sub ClsTT(A As Access.Application, TT$)
Dim T$, I
For Each I In TermAy(TT)
    T = I
    ClsTbl A, T
Next
End Sub
Sub ClsTbl(A As Access.Application, T)
DoCmd.Close acTable, T
End Sub

Sub ClsAllTbl(A As Access.Application)
Dim T As AccessObject
For Each T In A.CodeData.AllTables
    A.DoCmd.Close acTable, T.Name
Next
End Sub

Function ShwAcs(A As Access.Application) As Access.Application
If Not A.Visible Then A.Visible = True
Set ShwAcs = A
End Function

Function CvAcs(A) As Access.Application
Set CvAcs = A
End Function

Function Acs() As Access.Application
Static X As Access.Application: If IsNothing(X) Then Set X = New Access.Application
Set Acs = X
End Function

Sub CpyAcsFrm(A As Access.Application, FB)
Dim I As AccessObject
For Each I In A.CodeProject.AllForms
    A.DoCmd.CopyObject FB, , acForm, I.Name
Next
End Sub

Sub CpyAcsMd(A As Access.Application, ToFb)
Dim I As AccessObject
For Each I In A.CodeProject.AllModules
    A.DoCmd.CopyObject ToFb, , acModule, I.Name
Next
End Sub
Sub CpyAcsTbl(A As Access.Application, ToFb)

End Sub
Sub CpyAcsObj(A As Access.Application, ToFb)
Dim FB$
If HasFfn(ToFb) Then
    FB = NxtFfnzAva(A.CurrentDb.Name)
Else
    'Fb = Fb0
End If
Ass HasPth(Pth(FB))
Ass Not HasFfn(FB)
CrtFb FB
CpyAcsTbl A, FB
CpyAcsFrm A, FB
CpyAcsMd A, FB
CpyAcsRf A, FB
End Sub
Sub CpyAcsRf(A As Access.Application, ToFb)

End Sub
Sub SelPthzTxtb(A As Access.TextBox)
Dim R$
R = SelPth(A.Value)
If R = "" Then Exit Sub
A.Value = R
End Sub
Sub TurnOffTabStop(AcsCtl As Access.Control)
Dim A As Access.Control
Set A = AcsCtl
If Not HasPfx(A.Name, "Cmd") Then Exit Sub
Select Case True
Case IsBtn(A): CvBtn(A).TabStop = False
Case IsTglBtn(A): CvAcsTgl(A).TabStop = False
End Select
End Sub


'Assume there is Application.Forms("Main").Msg (TextBox)
'MMsg means Main.Msg (TextBox)
Sub ClrMainMsg()
SetMainMsg ""
End Sub

Sub SetMainMsgzQnm(QryNm)
SetMainMsg "Running query: (" & QryNm & ")...."
End Sub

Sub SetMainMsg(Msg$)
On Error Resume Next
SetTBox MMBox, Msg
End Sub

Private Property Get MMBox() As Access.TextBox
On Error Resume Next
Set MMBox = MFrm.Controls("Msg")
End Property

Private Property Get MFrm() As Access.Form
On Error Resume Next
Set MFrm = Access.Forms("Main")
End Property

Private Sub ZZ()
Dim A As Variant
Dim B$
ClrMainMsg
SetMainMsgzQnm A
SetMainMsg B
End Sub

Sub FrmSetCmdNotTabStop(A As Access.Form)
DoItrFun A.Controls, "CmdTurnOffTabStop"
End Sub

Function CvAcsCtl(A) As Access.Control
Set CvAcsCtl = A
End Function

Function CvAcsBtn(A) As Access.CommandButton
Set CvAcsBtn = A
End Function

Function CvAcsTgl(A) As Access.ToggleButton
Set CvAcsTgl = A
End Function

Sub SetTBox(A As Access.TextBox, Msg$)
Dim CrLf$, B$
If A.Value <> "" Then CrLf = vbCrLf
B = LasNLines(A.Value & CrLf & Now & " " & Msg, 5)
A.Value = B
DoEvents
End Sub

Function PjzFba(Fba, A As Access.Application) As VBProject
OpnFb A, Fba
Set PjzFba = PjzAcs(A)
End Function

Function PjzAcs(A As Access.Application) As VBProject
Set PjzAcs = A.Vbe.ActiveVBProject
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

Function NewAcs(Optional Shw As Boolean) As Access.Application
Dim O As Access.Application: Set O = CreateObject("Access.Application")
If Shw Then O.Visible = True
Set NewAcs = O
End Function

Function DbNmzAcs$(A As Access.Application)
On Error Resume Next
DbNmzAcs = A.CurrentDb.Name
End Function

Sub OpnFb(A As Access.Application, FB)
If DbNmzAcs(A) = FB Then Exit Sub
ClsDbzAcs A
A.OpenCurrentDatabase FB
End Sub

Function DftAcs(A As Access.Application) As Access.Application
If IsNothing(A) Then
    Set DftAcs = NewAcs
Else
    Set DftAcs = A
End If
End Function

'http://www.utteraccess.com/forum/USysRegInf-table-t353963.html
''able Name = USysRegInf
'Fields: Subkey (text), Type (number), ValName (text), Value (text)
'At least 3 records.
'Subkey in all 3 records = 'HKEY_CURRENT_ACCESS_PROFILE\Menu Add-Ins\&NameOfYourAdd-inHere'
'Type in 1st record = '0' then '1' in last 2 records
'ValName is blank in first record, then 'Expression' in second and 'Library' in the thid record.
'Value is blank in first record, then '=NameOfFunctionToOpenFormInYourDatabase()' in the second record and '|ACCDIR\NameOfYourDatabase.mde' in the third record.
'That is the best I can suggest. You may need more records depending on your Add-in. Do not add the single quotes (') in the description of what goes in each record.
'hth,
'Jac"
'SK = 'HKEY_CURRENT_ACCESS_PROFILE\Menu Add-Ins\&NameOfYourAdd-inHere
' Rec#  SubKey Type ValName        Value
' 1      SK    0     ""            ""
' 2      SK    1     "Expression"  "={FunNm}()"
' 3      SK    1     "Library"     "|ACCDIDR\{fba}"
Sub CrtTblzUSysRegInf(A As Database)
RunQ A, "Create Table [USysRegInf] (Subky Text,Type Long,ValName Text,Value Text)"
End Sub
Sub EnsTblzUSysRegInf(A As Database)
If HasTbl(A, "USysRegInf") Then CrtTblzUSysRegInf A
End Sub

Sub InstallAddin(A As Database, FB, Optional AutoFunNm$ = "AutoExec")
Dim Sk$: Sk = "HKEY_CURRENT_ACCESS_PROFILE\Menu Add-Ins\&NameOfYourAdd-inHere"
Dim Fba$: Fba = ""
Dim FunNm$
Stop '
RunQQ A, "Insert into [USysRegInf] Values('?',0,'','')"
RunQQ A, "Insert into [USysRegInf] Values('?',1,'Expression','?')", Sk, FunNm
RunQQ A, "Insert into [USysRegInf] Values('?',1,'Library','?')", Sk, Fba
End Sub

