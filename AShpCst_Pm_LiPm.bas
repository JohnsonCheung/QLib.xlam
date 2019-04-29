Attribute VB_Name = "AShpCst_Pm_LiPm"
Option Explicit
Private Db As Database
Private Property Get ShpCstLnkPmLy() As String()
Const LnkColVblzZHT1$ = _
    " ZHT1   D Brand  |" & _
    " RateSc M Amount |" & _
    " VdtFm  M [Valid From]  |" & _
    " VdtTo  M [Valid to]"

Const LnkColVblzUom$ = _
    "Sku    M Material |" & _
    "Des    M [Material Description] |" & _
    "Sc_U   M SC |" & _
    "StkUom M [Base Unit of Measure] |" & _
    "Topaz  M [Topaz Code] |" & _
    "ProdH  M [Product hierarchy]"
 
Const LnkColVblzMB52$ = _
    " Sku    M Material |" & _
    " Whs    M Plant    |" & _
    " QInsp  D [In Quality Insp#]|" & _
    " QUnRes D Unrestricted|" & _
    " QBlk   D Blocked"
'A = "MB52": PushObj O, LiFxLnkColVbl(A, A, "Sheet1", LnkColVblzMB52)
'A = "UOM":  PushObj O, LiFxLnkColVbl(A, A, "Sheet1", LnkColVblzUom)
'            PushObj O, LiFxLnkColVbl("ZHT1", "ZHT18701", "8701", LnkColVblzZHT1)
'            PushObj O, LiFxLnkColVbl("ZHT1", "ZHT18601", "8601", LnkColVblzZHT1)
End Property
