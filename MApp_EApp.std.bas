Attribute VB_Name = "MApp_EApp"
Option Explicit
Public Const AppHomMHD$ = "C:\Users\user\Desktop\MHD\SAPAccessReports\"
Public Const AppStkShpRateFb$ = AppHomMHD & "StockShipRate\StockShipRate\StockShipRate (ver 1.0).accdb"
Public Const AppTaxExpCmpFb$ = AppHomMHD & "TaxExpCmp\TaxExpCmp\TaxExpCmp v1.3.accdb"
Public Const AppStkShpCstFb$ = AppHomMHD & "StockShipCost\StockShipCost (ver 1.0).accdb"
Public Const AppTaxRateAlertFb$ = AppHomMHD & "TaxRateAlert\TaxRateAlert\TaxRateAlert (ver 1.3).accdb"
Public Const AppJJFb$ = AppHomMHD & "TaxExpCmp\TaxExpCmp\PgmObj\Lib\jj.accdb"

Property Get EAppFbDic() As Dictionary
Const A$ = "N:\SAPAccessReports\"
Set EAppFbDic = New Dictionary
With EAppFbDic
        .Add "Duty", A & "DutyPrepay\.accdb"
       .Add "SkHld", A & "StkHld\.accdb"
     .Add "ShpRate", A & "DutyPrepay\StockShipRate_Data.accdb"
      .Add "ShpCst", A & "StockShipCost\.accdb"
      .Add "TaxCmp", A & "TaxExpCmp\.accdb"
    .Add "TaxAlert", A & "TaxRateAlert\.accdb"
End With
End Property
