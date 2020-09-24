VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_PhTmp5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Private Sub B_Crt(): Crt: End Sub
Private Sub B_W1Add6ColAdd6Col()
Crt
DoCmd.RunSQL "Select Sku into [$PHLSku] from Sku"
W1Add6ColAdd6Col
BrwTC "#A"
Brw
End Sub
Private Sub B_W1Add3ColBus()
RunqC "Select Sku into [$FcSku] from Sku where WithOHHst"
W1Add3ColBus
BrwTC "$FcSku"
End Sub

Sub Drp(): DrpTTC "$PHL1 $PHL2 $PHL3 $PHL4 $PHLSku": End Sub

Sub Brw(): BrwTbQtp2C "$PH?", MH.PH.Ss5: End Sub
Sub Crt() ' #Create-(5PH)-Tables:PHL? 1 2 3 4 Sku#
StsQry "Creating 5 tables [$PHL? 1 2 3 4 Sku]"
'Build:
'    D StruQtpC("$PHL? Sku 1 2 3 4")
'    $PHL1   PHL1 Srt1 PHNam
'    $PHL2   PHL2 Srt2 PHNam PHBrd
'    $PHL3   PHL3 Srt3 PHNam PHBrd PHQGp
'    $PHL4   PHL4 Srt4 PHNam PHBrd PHQGp PHQly
'    $PHLSku Sku  PHL4 Srt4 PHNam PHBrd PHQGp PHQly SkuDes BusArea Bus
'From:
'    ProdHierarchy PH      | Lvl Des Sno Srt WithOHCur WithOHHst DteCrt DteUpd
'    D StruTTzC("PHLBus PHLStm ProdHierarchy Sku Topaz")
'    PHLBus        BusArea | PHBus PHSBus Stm DteCrt
'    PHLStm        Stm     | Stm2 Stream PHSStm
'    Sku           *       | WithOHHst WithOHCur Topaz *Des ProdHierarchy StkUnit Litre/Btl Btl/Ac Unit/Ac Unit/Sc BusArea TaxRateHK
'                          | TaxUOMHK TaxRateMO TaxUOMMO DteCrt DteRUpdTopaz DteRUpdTaxRate BusAreaSap BusAreaOvr Litre/BtlSap Litre/BtlOvr
'    Topaz         *       | Cd* DteCrt
'Note:
'  Only those WithOHHst will be put into $PH{5}
'Steps:
'   WTmpL1234 :: ProdHirarchy ->    ($PHL1..4)
'   WTmpLSku  ::                 +) (Sku Topaz PHLStm) $PHL1..4)
'                                 =>  $PHLSku
WL1234
W1LSku
End Sub
Private Sub WL1234()
'ProdHierarchy -> #PH -> (#L1..4) ->($PHL1..4)
'#PH :: Lvl PH   Des   Srt
'#L1 ::     PHL1 PHNam Srt1
'#L2 ::     PHL2 PHBrd Srt2
'#L3 ::     PHL3 PHQGp Srt3
'#L4 ::     PHL4 PHQly Srt4
'#L1       -> $PHL1 ::   PHL1   Srt1   PHNam
'(#L2 #L1) -> $PHL2 :: x.PHL2 x.Srt2 x.PHNam a.PHBrd
'(#L3 #L2) -> $PHL3 :: x.PHL3 x.Srt3 x.PHNam x.PHBrd a.PHQGp
'(#L4 #L3) -> $PHL4 :: x.PHL4 x.Srt4 x.PHNam x.PHBrd x.PHQGp a.PHQly
DrpPfxxC "#PH #L $PHL"
RunqC "Select Lvl,PH,Des,Srt Into [#PH] from ProdHierarchy where WithOHHst" 'Note: Only ProdHierarchy->WithOHHst = True
RunqC "SELECT Left(PH,2) AS PHL1,                                                     Des AS PHNam, Srt AS Srt1 Into [#L1] FROM [#PH] WHERE Lvl=1"
RunqC "SELECT Left(PH,2) AS PHL1, Left(PH,4) AS PHL2,                                 Des AS PHBrd, Srt AS Srt2 Into [#L2] FROM [#PH] WHERE Lvl=2"
RunqC "SELECT Left(PH,2) AS PHL1, Left(PH,4) AS PHL2, Left(PH,7) AS PHL3,             Des AS PHQGp, Srt AS Srt3 Into [#L3] FROM [#PH] WHERE Lvl=3"
RunqC "SELECT Left(PH,2) AS PHL1, Left(PH,4) AS PHL2, Left(PH,7) AS PHL3, PH AS PHL4, Des AS PHQly, Srt AS Srt4 Into [#L4] FROM [#PH] WHERE Lvl=4"
RunqC "SELECT PHL1,Srt1,PHNam                   INTO [$PHL1] FROM [#L1]                                        "
RunqC "Select PHL2,Srt2,PHNam,PHBrd             INTO [$PHL2] FROM [#L2] X inner join [$PHL1] a on x.PHL1=a.PHL1"
RunqC "Select PHL3,Srt3,PHNam,PHBrd,PHQGp       INTO [$PHL3] FROM [#L3] X inner join [$PHL2] a on x.PHL2=a.PHL2"
RunqC "Select PHL4,Srt4,PHNam,PHBrd,PHQGp,PHQly INTO [$PHL4] FROM [#L4] x inner join [$PHL3] a on x.PHL3=a.PHL3"
DrpPfxxC "#PH #L"
End Sub
Private Sub W1LSku()
DrpC "$PHLSku"

'$PHLSku 3 fields: Sku SkuDes Stream
RunqC "Select Sku,SkuDes,CdTopaz,'' As Stm,'' As Stream into [$PHLSku] from Sku x inner join Topaz a on x.Topaz=a.Topaz where WithOHHst"
RunqC "Update [$PHLSku] set Stm=IIf(Left(CdTopaz,3)='UDV','U','M')"
RunqC "Update [$PHLSku] x inner join PHLStm a on a.Stm=x.Stm set x.Stream=a.Stream"
RunqC "Alter Table [$PHLSku] Drop Column Stm,CdTopaz"
RunqC "Update [$PHLSku] x inner join [$PHLSku] a on a.Sku=x.Sku set x.Stream=a.Stream"

W1Add6ColAdd6Col   'Add 6 fields: PHL4 PHNam PHBrd PHQGp PHQly ($PHL5 must exist)"
W1Add3ColBus  'Add 3 fields: BusArea PHBus SrtBus
'Tot 12 fields: Sku SkuDes Stream | Srt4 PHL4 PHNam PHBrd PHQGp PHQly | BusArea PHBus SrtBus
End Sub

Private Sub W1Add6ColAdd6Col() 'Add 6 columns to $PHLSku.  6 columns: Srt4 PHL4 PHNam PHBrd PHQGp PHQly
RunqC "Alter Table [$PHLSku] add column Srt4 Text(10), PHL4 Text(10), PHNam Text(255), PHBrd Text(255), PHQGp Text(255), PHQly Text(255)"
RunqC "Update [$PHLSku] x inner join Sku a on a.Sku=x.Sku set x.PHL4=Left(ProdHierarchy,10)"
RunqC "Update [$PHLSku] x inner join [$PHL4] a on a.PHL4=x.PHL4 set x.Srt4=a.Srt4,x.PHNam=a.PHNam,x.PHBrd=a.PHBrd,x.PHQGp=a.PHQGp,x.PHQly=a.PHQly"
End Sub
Private Sub W1Add3ColBus() ' Add 3 columns to $PHLSku.  3 columns: BusArea PHBus SrtBus
RunqC "Alter Table [$PHLSku] add column BusArea Text(10), PHBus Text(255), SrtBus Byte"
RunqC "Update [$PHLSku] x inner join Sku a on a.Sku=x.Sku set x.BusArea=a.BusArea"
RunqC "Update [$PHLSku] x inner join PHLBus a on a.BusArea=x.BusArea set x.PHBus=a.PHBus,x.SrtBus=PHSBus"
End Sub
