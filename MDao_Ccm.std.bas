Attribute VB_Name = "MDao_Ccm"
Option Explicit
Const CMod$ = "MDao__Ccm."
Private Sub Z_LnkCcm()
Dim D As Database, IsLcl As Boolean
Set D = Db(SampFbzShpRate)
IsLcl = True
GoSub Tst
Exit Sub
Tst:
    LnkCcm D, IsLcl
    Return
End Sub
Sub LnkCcm(Db As Database, IsLcl As Boolean)
'Ccm stands for Space-[C]ir[c]umflex-accent
'CcmTbl is ^xxx table in Db (pgm-database),
'          which should be same stru as N:\..._Data.accdb @ xxx
'          and   data should be copied from N:\..._Data.accdb for development purpose
'At the same time, in Db, there will be xxx as linked table either
'  1. In production, linking to N:\..._Data.accdb @ xxx
'  2. In development, linking to Db @ ^xxx
'Notes:
'  The TarFb (N:\..._Data.accdb) of each CcmTbl may be diff
'      They are stored in Description of CcmTbl manual, it is edited manually during development.
'  those xxx table in Db will be used in the program.
'  and ^xxx is create manually in development and should be deployed to N:\..._Data.accdb
'  assume Db always have some ^xxx, otherwise throw
'This Sub is to re-link the xxx in given [Db] to
'  1. [Db] if [TarFb] is not given
'  2. [TarFb] if [TarFb] is given.
Const CSub$ = CMod & "LnkCcm"
Dim T$()  ' All ^xxx
    T = CcmTny(Db)
    If Sz(T) = 0 Then Thw CSub, "No ^xxx table in [Db]", Db.Name 'Assume always
Chk Db, T, IsLcl ' Chk if all T after rmv ^ is in TarFb
LnkCcmzTny Db, T, IsLcl
End Sub

Private Sub LnkCcmzTny(Db As Database, CcmTny$(), IsLcl As Boolean)
Dim CcmTbl, TarFb$
TarFb = Db.Name
For Each CcmTbl In CcmTny
    If FstChr(CcmTbl) <> "^" Then Thw CSub, "All table in CcmTny must begin ^", "Tbl-without-^ CcmTny", CcmTbl, CcmTny
    LnkFb Db, RmvFstChr(CcmTbl), TarFb, CcmTbl
Next
End Sub
Private Sub Chk(Db As Database, CcmTny$(), IsLcl As Boolean)
Const CSub$ = CMod & "Chk"
If Not IsLcl Then Chk2 Db, CcmTny: Exit Sub ' Asserting for TarFb is stored in CcmTny's description

'Asserting for TarFb = Db
Dim Mis$(): Mis = Chk1(Db, CcmTny)
If Sz(Mis) = 0 Then Exit Sub
Thw CSub, "[Some-Missing-Tar-Tbl] in [Db] cannot be found according to given [CcmTny] in [Db]", Mis, Db.Name, CcmTny, Db.Name
End Sub
Private Function Chk1(A As Database, CcmTny$()) As String()
Dim N1$(): N1 = Tny(A)
Dim N2$(): N2 = AyRmvFstChr(CcmTny)
Chk1 = AyMinus(N2, N1)
End Function

Private Sub Chk2(Db As Database, CcmTny$())
'Throw if any Corresponding-Table in TarFb is not found
Dim O$(), T
For Each T In CcmTny
    PushIAy O, Chk3(Db, T)
Next
'ErThw O
End Sub
Private Function Chk3(Db As Database, CcmTbl) As String()
Dim TarFb$
'    TarFb = Dbt_Des(Db, CcmTbl)
Select Case True
Case TarFb = "":            Chk3 = LyzMsgNap("[CcmTbl] in [Db] should have 'Des' which is TarFb, but this TarFb is blank", CcmTbl, Db.Name)
'Case NotHasFfn(TarFb):    Chk3 = LyzMsgNap("[CcmTbl] in [Db] should have [Des] which is TarFb, but this TarFb does not Has", CcmTbl, Db.name, TarFb)
Case Not HasFbt(TarFb, RmvFstChr(CcmTbl)):
    Chk3 = LyzMsgNap("[CcmTbl] in [Db] should have [Des] which is TarFb, but this TarFb does not Has [Tbl-RmvFstChr(CcmTbl)]", CcmTbl, Db.Name, TarFb, RmvFstChr(CcmTbl))
End Select
End Function

Private Function CcmTny(Db As Database) As String()
CcmTny = AywPfx(Tny(Db), "^")
End Function

Private Sub Z_CcmTny()
Dim D As Database
'
Set D = Db(SampFbzShpRate)
Ept = SySsl("^CurYM ^IniRate ^IniRateH ^InvD ^InvH ^YM ^YMGR ^YMGRnoIR ^YMOH ^YMRate")
GoSub Tst
Exit Sub
Tst:
    Act = CcmTny(D)
    C
    Return
End Sub

Private Sub Z()
Z_LnkCcm
Z_CcmTny
MDao__Ccm:
End Sub
