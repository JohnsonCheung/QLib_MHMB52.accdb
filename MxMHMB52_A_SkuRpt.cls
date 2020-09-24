VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_SkuRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const SkuIoTp$ = "C:\Users\Public\Logistic\StockHolding8\WorkingDir\Templates\Sku List Template.xlsx"
Private WWPm As MxMHMB52_A_SkuPm
Private WWIo As New MxMHMB52_A_SkuIO
Sub SkuRpt()
W_OupSku_1

Dim Fxo$: Fxo = WWIo.Fxo
ExpAttTpC "SkuTp.xlsx", Fxo
Dim B As Workbook: Set B = WbFx(Fxo)
RfhWb B, CFb
DrpC "@Sku"

If WWPm.IsCpyTo Then
    Dim P$: P = WWPm.PthCpyTo
    Fso.CopyFile Fxo, P, True
End If
End Sub
Private Sub W_OupSku_1()
MH.TbSku.RfhFldWithOHXxx
W1_TmpQly_2
W1_TmpSkuMain
Const QSku$ = "SELECT x.SKU, x.SkuDes, x.WithOHHst, x.WithOHCur, x.CdTopaz, x.ProdHierarchy AS PH, x.StkUnit," & _
" x.[Litre/Btl], x.[Unit/AC], x.[Unit/Sc], x.[Btl/AC], x.[Btl/SC], x.[Btl/Unit]," & _
" x.TaxRateHK, x.TaxUOMHK, x.TaxRateMO, x.TaxUOMMO, x.DteCrt, x.DteRUpdTopaz AS DteUpd, x.DteRUpdTaxRate," & _
" a.Srt4, c.PHSStm, b.PHSBus," & _
" x.Stm, a.PHL1, a.PHL2, a.PHL3, x.PHL4, x.BusArea, x.Stream, a.PHNam, a.PHBrd, a.PHQGp, a.PHQly, b.PHBus" & _
" Into [@SKU]" & _
" FROM (([#SkuMain] AS x" & _
" LEFT JOIN [#Qly] AS a ON x.PHL4 = a.PHL4)" & _
" LEFT JOIN PHLBus As b ON x.BusArea = b.BusArea)" & _
" LEFT JOIN PHLStm As c ON x.Stm = c.Stm;"
DrpC "@Sku"
RunqC QSku
DrpTTC "#Qly #SkuMain"
'BrwTC "@Sku":Stop
End Sub
Private Sub W1_TmpQly_2()
W2_TmpL1234
DrpC "#Qly"
RunqC "SELECT x.*, L1.PHNam, L2.PHBrd, L3.PHQGp" & _
" INTO [#Qly]" & _
" FROM (([#L4] x" & _
" INNER JOIN [#L1] as L1 ON x.PHL1 = L1.PHL1)" & _
" INNER JOIN [#L2] as L2 ON x.PHL2 = L2.PHL2)" & _
" INNER JOIN [#L3] as L3 ON x.PHL3 = L3.PHL3;"
DrpQtpC "#L? 1 2 3 4"
End Sub
Private Sub W2_TmpL1234()
DrpTTC "#L1 #L2 #L3 #L4"
Const QL1$ = "SELECT Left([PH],2) AS PHL1, x.Des AS PHNam, x.Srt AS Srt1, x.WithOHCur, x.WithOHHst" & _
" INTO [#L1]" & _
" FROM ProdHierarchy x" & _
" WHERE (((x.Lvl)=1));"

Const QL2$ = "SELECT Left([PH],2) AS PHL1, Left([PH],4) AS PHL2, x.Des AS PHBrd, x.Srt AS Srt2, x.WithOHCur, x.WithOHHst" & _
" INTO [#L2]" & _
" FROM ProdHierarchy x" & _
" WHERE (((x.Lvl)=2));"

Const QL3$ = "SELECT Left([PH],2) AS PHL1, Left([PH],4) AS PHL2, Left([PH],7) AS PHL3, x.Des AS PHQGp, x.Srt AS Srt3, x.WithOHCur, x.WithOHHst" & _
" INTO [#L3]" & _
" FROM ProdHierarchy x" & _
" WHERE (((x.Lvl)=3));"

Const QL4$ = "SELECT Left([PH],2) AS PHL1, Left([PH],4) AS PHL2, Left([PH],7) AS PHL3, x.PH AS PHL4, x.Des AS PHQly, x.Srt AS Srt4, x.WithOHCur, x.WithOHHst" & _
" INTO [#L4]" & _
" FROM ProdHierarchy x" & _
" WHERE (((x.Lvl)=4));"
RunqSqyC Sy(QL1, QL2, QL3, QL4)
End Sub
Private Sub W1_TmpSkuMain()
Const QMain$ = "SELECT" & _
" IIf(Left([CdTopaz],3)='UDV','U','M') AS Stm," & _
" IIf(Left([CdTopaz],3)='UDV','Diageo','MH') AS Stream," & _
" a.CdTopaz, x.*," & _
" Left([ProdHierarchy],10) AS PHL4," & _
" IIf([Unit/AC]=0,0,[Btl/AC]/[Unit/AC]*[Unit/SC]) AS [Btl/SC]," & _
" IIf(Nz([Unit/AC],0)=0,Null,[Btl/Ac]/[Unit/Ac]) AS [Btl/Unit]" & _
" Into [#SkuMain]" & _
" FROM SKU as x" & _
" INNER JOIN Topaz as a ON a.Topaz = x.Topaz;"
DrpC "#SkuMain"
RunqC QMain
End Sub


Private Sub Class_Initialize()
Set WWPm = MH.SkuPm
End Sub
