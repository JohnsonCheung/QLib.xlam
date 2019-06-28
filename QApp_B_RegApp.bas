Attribute VB_Name = "QApp_B_RegApp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MApp_EApp."
Private Const Asm$ = "QApp"
Const H$ = "C:\Users\user\Desktop\MHD\SAPAccessReports\"
Public Const AppStkShpRateFb$ = H & "StockShipRate\StockShipRate\StockShipRate (ver 1.0).accdb"
Public Const AppTaxExpCmpFb$ = H & "TaxExpCmp\TaxExpCmp\TaxExpCmp v1.3.accdb"
Public Const AppStkShpCstFb$ = H & "StockShipCost\StockShipCost (ver 1.0).accdb"
Public Const AppTaxRateAlertFb$ = H & "TaxRateAlert\TaxRateAlert\TaxRateAlert (ver 1.3).accdb"
Public Const AppJJFb$ = H & "TaxExpCmp\TaxExpCmp\PgmObj\Lib\jj.accdb"

Property Get MHDAppFbDic() As Dictionary
Const A$ = "N:\SAPAccessReports\"
Erase XX
X "Duty     " & A & "DutyPrepay\.accdb"
X "SkHld    " & A & "StkHld\.accdb"
X "ShpRate  " & A & "DutyPrepay\StockShipRate_Data.accdb"
X "ShpCst   " & A & "StockShipCost\.accdb"
X "TaxCmp   " & A & "TaxExpCmp\.accdb"
X "TaxAlert " & A & "TaxRateAlert\.accdb"
Set MHDAppFbDic = Dic(XX)
Erase XX
End Property

Function AppFbAy() As String()
PushI AppFbAy, AppJJFb
PushI AppFbAy, AppStkShpCstFb
PushI AppFbAy, AppStkShpRateFb
PushI AppFbAy, AppTaxExpCmpFb
PushI AppFbAy, AppTaxRateAlertFb
End Function

