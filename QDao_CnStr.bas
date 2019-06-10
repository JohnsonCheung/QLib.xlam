Attribute VB_Name = "QDao_CnStr"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_CnStr."
Private Const Asm$ = "QDao"
Function CnStrzFbDao$(A)
CnStrzFbDao = ";DATABASE=" & A & ";"
End Function

Function CnStrzFxDao$(Fx)
'Excel 8.0;HDR=YES;IMEX=2;DATABASE=D:\Data\MyDoc\Development\ISS\Imports\PO\PUR904 (On-Line).xls;TABLE='PUR904 (On-Line)'
'INTO [Excel 8.0;HDR=YES;IMEX=2;DATABASE={0}].{1} FROM {2}"
'Excel 12.0 Xml;HDR=YES;IMEX=2;ACCDB=YES;DATABASE=C:\Users\sium\Desktop\TaxRate\sales text.xlsx;TABLE=Sheet1$
Dim O$
Select Case LCase(Ext(Fx))
Case ".xlsx":: O = "Excel 12.0 Xml;HDR=YES;IMEX=2;ACCDB=YES;DATABASE=" & Fx & ";"
Case ".xls": O = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=" & Fx & ";"
Case Else: Stop
End Select
CnStrzFxDao = O
End Function

Function CnStrzFbzAsAdoOle$(FB) 'Return a connection used as WbConnection
CnStrzFbzAsAdoOle = "OLEDb;" & CnStrzFbzAsAdo(FB)
End Function

Function CnStrzFbzForWc$(FB)
CnStrzFbzForWc = CnStrzFbzAsAdoOle(FB)
'CnStrzFbzAsForWbCn = FmtQQ("Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share Deny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserINF Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False", A)
'CnStrzFbzAsForWbCn = FmtQQ("OLEDB;Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share Deny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserINF Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False", A)
End Function

