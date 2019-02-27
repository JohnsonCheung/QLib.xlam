Attribute VB_Name = "MAcs"
Option Explicit
Const CMod$ = "MAcs."
Function CAcs() As Access.Application
Static A As New Access.Application
OpnFbz A, CDb.Name
End Function
Sub SavRec()
DoCmd.RunCommand acCmdSaveRecord
End Sub
Function AcsFb$(A As Access.Application)
On Error Resume Next
AcsFb = A.CurrentDb.Name
End Function

Sub ClsDb(A As Access.Application)
On Error Resume Next
A.CloseCurrentDatabase
End Sub
Sub OpnFbz(A As Access.Application, Fb)
If FbzAcs(A) = Fb Then Exit Sub
ClsDb A
A.OpenCurrentDatabase Fb
End Sub

Sub BrwFb(Fb)
Static Acs As New Access.Application
OpnFbz Acs, Fb
Acs.Visible = True
End Sub

Sub ClsTTz(A As Access.Application, TT)
Dim T
For Each T In FnyzFF(TT)
    ClsTblz A, T
Next
End Sub
Sub ClsTblz(A As Access.Application, T)
DoCmd.Close acTable, T
End Sub

Sub ClsAllTblz(A As Access.Application)
Dim T As AccessObject
For Each T In A.CodeData.AllTables
    A.DoCmd.Close acTable, T.Name
Next
End Sub

Function FbzAcs$(A As Access.Application)
FbzAcs = AcsFb(A)
End Function

Sub QuitzA(A As Access.Application)
ClsDb A
A.Quit
Set A = Nothing
End Sub

Function AcsVis(A As Access.Application) As Access.Application
If Not A.Visible Then A.Visible = True
Set AcsVis = A
End Function

Function CvAcs(A) As Access.Application
Set CvAcs = A
End Function

Property Get Acs() As Access.Application
Set Acs = Access.Application
End Property

Sub CpyAllAcsFrm(A As Access.Application, Fb$)
Dim I As AccessObject
For Each I In A.CodeProject.AllForms
    A.DoCmd.CopyObject Fb, , acForm, I.Name
Next
End Sub

Sub CpyAcsMd(A As Access.Application, ToFb$)
Dim I As AccessObject
For Each I In A.CodeProject.AllModules
    A.DoCmd.CopyObject ToFb, , acModule, I.Name
Next
End Sub

Sub CpyAcsObj(A As Access.Application, ToFb$)
Dim Fb$
If HasFfn(ToFb) Then
    Fb = NxtFfn(A.CurrentDb.Name)
Else
    'Fb = Fb0
End If
Ass HasPth(Pth(Fb))
Ass Not HasFfn(Fb)
CrtFb Fb
'AcsCpyTbl A, Fb
'AcsCpyFrm A, Fb
'AcsCpyMd A, Fb
'AcsCpyRf A, Fb
End Sub

Sub TxtbSelPth(A As Access.TextBox)
Dim R$
R = PthSel(A.Value)
If R = "" Then Exit Sub
A.Value = R
End Sub
Sub CmdTurnOffTabStop(AcsCtl As Access.Control)
Dim A As Access.Control
Set A = AcsCtl
If Not HasPfx(A.Name, "Cmd") Then Exit Sub
Select Case True
Case IsBtn(A): CvBtn(A).TabStop = False
Case IsTgl(A): CvTgl(A).TabStop = False
End Select
End Sub


'Assume there is Application.Forms("Main").Msg (TextBox)
'MMsg means Main.Msg (TextBox)
Sub ClrMainMsg()
SetMainMsg ""
End Sub

Sub SetMainMsgzQnm(QryNm)
SetMainMsg "Running query: " & QryNm
End Sub

Sub SetMainMsg(A$)
On Error Resume Next
SetTBox MMBox, A
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

Function CvCtl(A) As Access.Control
Set CvCtl = A
End Function

Function CvBtn(A) As Access.CommandButton
Set CvBtn = A
End Function

Function CvTgl(A) As Access.ToggleButton
Set CvTgl = A
End Function

Sub SetTBox(A As Access.TextBox, Msg$)
Dim CrLf$, B$
If A.Value <> "" Then CrLf = vbCrLf
B = LasNLines(A.Value & CrLf & Now & " " & Msg, 5)
A.Value = B
DoEvents
End Sub

Sub AcsQuit(A As Access.Application)
On Error Resume Next
Stamp "AcsQuit: Begin"
Stamp "AcsQuit: Cls":         A.CloseCurrentDatabase
Stamp "AcsQuit: Quit":        A.Quit
Stamp "AcsQuit: Set Nothing": Set A = Nothing
Stamp "AcsQuit: End"
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
Sub OpnFb(A As Access.Application, Fb)
If DbNmzAcs(A) = Fb Then Exit Sub
ClsDb A
A.OpenCurrentDatabase Fb
End Sub
Function DftAcs(A As Access.Application) As Access.Application
If IsNothing(A) Then
    Set DftAcs = NewAcs
Else
    Set DftAcs = A
End If
End Function

