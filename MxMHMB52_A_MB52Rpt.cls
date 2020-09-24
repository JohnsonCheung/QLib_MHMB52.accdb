VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_MB52Rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_MB52Rpt."

Property Get SsSapUnit$(): SsSapUnit = "Litre/Btl Btl/AC Unit/AC Unit/SC": End Property
Private Sub B_LasRpt():    RptLas:                                         End Sub
Private Sub B_WO_RatPac()
ClsTblAllC
WO_RatPac MH.TbOH.YmdLas
DoCmd.OpenTable "@PacRatD"
DoCmd.OpenTable "@PacRat"
End Sub
Private Sub B_WO_RatBchHK__BH()
Dim A As Ymd: A = MH.TbOH.YmdLas
ClsTblAllC
WO_RatPac A
WO_RatBchHK__BH A
DoCmd.OpenTable "@BchRatHK"
End Sub
Private Sub B_WO_RatBchMO()
WO_RatBchMO Ymd(20, 1, 30)
DoCmd.OpenTable "@BchRatMO"
End Sub
Private Sub B_WO_OH__OH()
Dim A As Ymd: A = Ymd(19, 11, 29)
ClsTblAllC
WO_RatZHT0
WO_RatBchHK__BH A '@BchRatHK: Using Sku BchNo BchRat
WO_RatBchMO A '@BchRatMO: Using Sku BchNo BchRat
WO_OH__OH A
DoCmd.OpenTable "@OH"
End Sub
Private Sub B_WBH_HshRatBch86()
Dim A As Ymd: A = MH.TbOH.YmdLas
ClsTblAllC
WBH_HshRatBch86_3p A
WBH_HshRatBch86_OH A
WBH_HshRatBch86 A
DoCmd.OpenTable "#RatBch86"
End Sub
Private Sub B_WBH_HshRatLas86()
Dim A As Ymd: A = MH.TbOH.YmdLas
ClsTblAllC
WBH_HshRatBch86_OH A
WBH_HshRatBch86 A
WBH_HshRatLas86 A
DoCmd.OpenTable "#RatLas86"
End Sub
Private Sub AA()
Dim F$: F = FxTmp: ExpAttTpC MH.MB52Tp.Fn, MH.MB52Tp.Tp
Dim S As Worksheet: Set S = WbFx(F).Sheets("Data") ' Data
Maxv S.Application
SetLoFmllny LoFst(S), MH.MB52IO.FmllnyWsDta
End Sub
Sub RptLas():         RptYmd MH.TbOH.YmdLas: End Sub
Sub RptYmd(A As Ymd): WR__Rpt A:             End Sub
Private Sub WR__Rpt(A As Ymd)
Dim Fxo$: Fxo = MH.MB52IO.Fxo(A)
If AskOpnFx(Fxo) Then Exit Sub
WR_DrpOup
WR_Oup__O A
WR_GenFx Fxo
WR_UpdTbReportFldDteGen A
WR_DrpOup
End Sub
Private Sub WR_Oup__O(A As Ymd)
WO_RatZHT0
WO_RatPac A
WO_RatBchHK__BH A
WO_RatBchMO A
WO_OH__OH A
End Sub
Private Sub WR_DrpOup(): DrpTTC "@BchRatHK @BchRatMO @OH @PacRat @PacRatD @RatZHT0": End Sub

Private Sub WR_UpdTbReportFldDteGen(A As Ymd)
RunqC "Update Report Set DteGen=#" & Now & "# " & MH.TbOH.WhYmd(A)
End Sub

Private Sub WO_OH__OH(A As Ymd)
'Aim: Create @OH from OH & GITDet & Add Atr-fields
'Inp: OH     = YY MM DD Co Sku YpStk BchNo | Btl Val
'Inp: GITDet = YY MM DD Co Sku             | Btl HKD
'Ref: Sku    : ZHT0-Rate  = Sku->TaxRate    (/AC)
'     SkuB   : Sku BchNo | DutyRateB        ' DutyRateB is /Btl rate
'     YpStk  = YpStk | NmYpStk Co SLoc YpCls IsTaxLoc
' ## Stp   Oup
'Inp: $OH =
'Oup: @OH =
'Why: Just Cpy $OH by order and Rename fields
DoCmd.SetWarnings False
StsQry "@OH":
Dim Git%: Git = MH.GitIO.YpStk
Dim Wh$: Wh = MH.TbOH.WhYmd(A)

'Crt $OH  ! OH
'Ins $OH  ! GIT
RunqC "Select" & _
                        " YY,MM,DD,Co,YpStk,SLoc,Sku,BchNo,Btl,Val Into [$OH] From OH" & Wh
RunqC "Insert into [$OH] (YY,MM,DD,Co,YpStk,     Sku,Btl,Val) Select YY,MM,DD,Co," & Git & ",Sku,Btl,HKD from GITDet" & Wh

'DrsAddDc $OH ! ..
RunqC "Alter Table [$OH] add column " & _
"NmYpStk Text(50),TaxLoc Text(1),[3p] Text(1), NoTax Text(1)," & _
"SkuDes Text(255),PH Text(20), CdTopaz Text(20), Topaz Long," & _
"[Litre/Btl] Double," & _
"[Btl/AC] Integer  ," & _
"[Unit/AC] Integer ," & _
"[Unit/SC] Double  ," & _
"StkUnit Text(3)   , " & _
"BusArea Text(40), PHL1 Text(2),PHL2 Text(4),PHL3 Text(7),PHL4 Text(10)," & _
"Stream Text(6),PHBus Text(99), PHNam Text(99),PHBrd Text(99),PHQGp Text(99),PHQly Text(99)," & _
"PHSStm Integer, PHSBus Integer, PHSrt1 Text(2), PHSrt2 Text(4), PHSrt3 Text(6), PHSrt4 Text(8)," & _
"BchRat Double, BchRatTy Text(4), RatZHT0 Double"
Const UpdJn$ = "Update [$OH] x inner join "
Const UpdSet$ = "Update [$OH] Set "

'UpdFld $OH->YpStk TaxLoc
'UpdFld $OH->3p
'UpdFld $OH->NoTax
RunqC UpdJn & "YpStk a on a.YpStk=x.YpStk        set x.NmYpStk = a.NmYpStk, x.TaxLoc=IIf(a.IsTaxLoc,'Y','')"
RunqC UpdJn & "SkuTaxBy3rdParty a on a.Sku=x.Sku set x.3p='Y'"
RunqC UpdJn & "SkuNoLongerTax a on a.Sku=x.Sku   set x.NoTax='Y'"

'UpdFld $OH->PH SkuDes Topaz ...
RunqC UpdJn & "Sku   a on a.Sku=x.Sku" & _
" set " & _
"x.PH=a.ProdHierarchy, x.SkuDes=a.SkuDes,x.Topaz=a.Topaz," & _
"x.[Litre/Btl]=a.[Litre/Btl]," & _
"x.[Btl/AC]   =a.[Btl/AC]   ," & _
"x.[Unit/AC]  =a.[Unit/AC]  ," & _
"x.[Unit/SC]  =a.[Unit/SC]  ," & _
"x.StkUnit    =a.StkUnit    ," & _
"x.BusArea    =a.BusArea"

RunqC UpdJn & "Topaz a on a.Topaz=x.Topaz set x.CdTopaz=a.CdTopaz"
RunqC UpdSet & "Stream=IIf(Left(CdTopaz,3)='UDV','Diageo','MH')"
RunqC "Alter Table [$OH] Drop Column Topaz"

'Stp-PH ===========================================================================================
'Set PH1..4                     $OH
'Set PHNam..Qly PHSrt1..4       #PH?
'Set PHBus PHSBus               PHLBus
'Set PHSStm                     PHLStm
RunqC UpdSet & "PHL1 = Left(PH,2), PHL2 = Left(PH,4), PHL3 = Left(PH,7), PHL4 = Left(PH,10)"

RunqC "Select Left(PH,2) as PHL1,  Des as PHNam, Srt As PHSrt1 Into [#PHL1] from ProdHierarchy where Lvl=1"
RunqC "Select Left(PH,4) as PHL2,  Des as PHBrd, Srt As PHSrt2 Into [#PHL2] from ProdHierarchy where Lvl=2"
RunqC "Select Left(PH,7) as PHL3,  Des as PHQGp, Srt As PHSrt3 Into [#PHL3] from ProdHierarchy where Lvl=3"
RunqC "Select Left(PH,10) as PHL4, Des as PHQly, Srt As PHSrt4 Into [#PHL4] from ProdHierarchy where Lvl=4"

RunqC UpdJn & "[#PHL1] a on x.PHL1=a.PHL1 set x.PHNam=a.PHNam, x.PHSrt1=a.PHSrt1"
RunqC UpdJn & "[#PHL2] a on x.PHL2=a.PHL2 set x.PHBrd=a.PHBrd, x.PHSrt2=a.PHSrt2"
RunqC UpdJn & "[#PHL3] a on x.PHL3=a.PHL3 set x.PHQGp=a.PHQGp, x.PHSrt3=a.PHSrt3"
RunqC UpdJn & "[#PHL4] a on x.PHL4=a.PHL4 set x.PHQly=a.PHQly, x.PHSrt4=a.PHSrt4"

DrpTTC JnPfxSno("#PHL", 1, 4)

RunqC UpdJn & "PHLStm a on a.Stream=x.Stream   Set x.PHSStm=a.PHSStm"
RunqC UpdJn & "PHLBus a on a.BusArea=x.BusArea Set x.PHBus=a.PHBus, x.PHSBus=a.PHSBus"

'DrsAddDc $OH->BchRat BchRatTy Co86
'DrsAddDc $OH->BchRat BchRatTy Co87
RunqC UpdJn & "[@BchRatHK] a on a.Sku=x.Sku and a.BchNo=x.BchNo Set x.BchRat=a.BchRat,x.BchRatTy=a.BchRatTy Where x.Co=86 and TaxLoc='Y'"
RunqC UpdJn & "[@BchRatMO] a on a.Sku=x.Sku and a.BchNo=x.BchNo Set x.BchRat=a.BchRat,x.BchRatTy='*MO'      Where x.Co=87 and TaxLoc='Y'"
RunqC UpdJn & "[@RatZHT0] a on x.Sku=a.Sku AND x.Co=a.Co set x.RatZHT0=a.RatZHT0 where TaxLoc='Y'"
'AddFld $OH-> All Fml Fld
Dim F: For Each F In Tm1y(MH.MB52IO.FmllnyWsDta)
    RunqC "Alter Table [$OH] Add Column [" & F & "] Byte"
Next
CrtTboRseq "OH", WOH_Fny, "Co,PHSStm,PHSrt4,[Litre/Btl],Sku"
End Sub
Private Function WOH_Fny() As String()
Dim O$()
PushSs O, "Co YY MM DD"
PushSs O, "Stream PHBus BusArea NmYpStk YpStk"
PushSs O, "SkuDes Sku"

PushSs O, "Litre/Btl Btl/AC Unit/SC Unit/AC StkUnit"
PushSs O, "Btl/Unit"
PushSs O, "ml/Btl Btl/AC' Litre/SC"
PushSs O, "Val Btl AC SC"

PushSs O, "BtlUPr"
PushSs O, "AcUPr"
PushSs O, "ScUPr"

PushSs O, "SLoc TaxLoc TaxItm NoTax 3p Taxed"
PushSs O, "BchRat RatZHT0"
PushSs O, "BchAmt ZHT0Amt"
PushSs O, "RatDif AmtDif"

PushSs O, "BchNo BchRatTy"

PushSs O, "CdTopaz PH"
PushSs O, "PHNam PHBrd PHQGp PHQly"
PushSs O, "PHSStm PHSBus PHSrt1 PHSrt2 PHSrt3 PHSrt4"
PushSs O, "PHL1 PHL2 PHL3 PHL4"
WOH_Fny = O
End Function

Private Sub WO_RatBchHK__BH(A As Ymd)
'Crt @BchRatHK : *Key *Rat *Bch *Las FmSkuCnt
'                *Key Co Sku Des
'                *Rat BchRatTy BchRat
'                *Las LasPermitD LasPermit LasPermitDate LasBchNo
'                *Bch BchPermitD BchPermit BchPermitDate BchNo
'   From  [SkuTaxBy3rdParty]  Sku RateU
'   From  [SkuNoLongerTax]    Sku
'   From  [#OHBch86]  Sku BchNo      ! From OH where Co=86 & Btl>0
'   From  [@PacRat] NewSku Rate                                      ! Rate is /Btl
'   From  [#RatBch86] Sku BchNo Rate PermitD Permit PermitDate  ! Rate is /Btl
'   From  [#RatLas86] Sku BchNo Rate PermitD Permit PermitDate  ! Rate is /Btl

StsQry "@BchRatHK"
DoCmd.SetWarnings False
DrpC "@BchRatHK"
WBH_HshRatBch86_OH A
WBH_HshRatBch86 A
WBH_HshRatLas86 A
Stop
'Crt @BchRatHK : empty
RunqC "Create Table [@BchRatHK] (Co Byte, BchRatTy Text(6), BchRat Currency, Sku Text(15), Des Text(255),BchNo Text(10)," & _
" BchPermitD Long, BchPermit Long, BchPermitDate Date," & _
" LasPermitD Long, LasPermit Long, LasPermitDate date, LasBchNo Text(10)," & _
" FmSkuCnt Byte, BtlRat Currency, [Btl/Ac] byte)"

'Ins @BchRatHK : for all OH
RunqC "Insert into [@BchRatHK] (Co,Sku,BchNo) select 86,Sku,BchNo from [#OHBch86]"
'Upd @BchRatHK : for *3p
'Upd @BchRatHK : for *NoTax
'Upd @BchRatHK : for *Bch
'Upd @BchRatHK : for *Las
'Upd @BchRatHK : for *Pac
RunqC "Update [@BchRatHK] x inner join [#3p]            a on x.Sku=a.Sku              set BchRatTy='*3p' ,x.BtlRat=a.BtlRat"
RunqC "Update [@BchRatHK] x inner join [SkuNoLongerTax] a on x.Sku=a.Sku              set BchRatTy='*NoTax'"
RunqC "Update [@BchRatHK] x inner join [#RatBch86] a on x.Sku=a.Sku and x.BchNo=a.BchNo set BchRatTy='*Bch',x.BtlRat=a.Rate,BchPermitD=PermitD,BchPermit=Permit,BchPermitDate=PermitDate"
RunqC "Update [@BchRatHK] x inner join [#RatLas86] a on x.Sku=a.Sku                     set BchRatTy='*Las',x.BtlRat=a.Rate,LasPermitD=PermitD,LasPermit=Permit,LasPermitDate=PermitDate,LasBchNo=a.BchNo where BchRatTy is null"
RunqC "Update [@BchRatHK] x inner join [@PacRat] a on x.Sku=a.NewSku                  set BchRatTy='*Pac',x.BtlRat=a.Rate,x.FmSkuCnt=a.FmSkuCnt"

'Upd @BchRatHK : ->[Btl/Ac], Des
'Dlt @BchRatHK : for no BchRatTy
'Upd @BchRatHK : ->[BchRat]
RunqC "Update [@BchRatHK] x inner join Sku a on x.Sku=a.Sku set x.[Btl/AC]=a.[Btl/Ac],x.Des=a.SkuDes"
RunqC "Delete * from [@BchRatHK] where BchRatTy is null or Nz([Btl/Ac],0)=0"
RunqC "Update [@BchRatHK] set BchRat=BtlRat*[Btl/Ac]"

'DrsDrpDc @BchRatHK
DrpTTC "#OHBch86 #RatBch86 #RatLas86 #3P"
End Sub
Private Sub WBH_HshRatBch86_3p(A As Ymd)
'Crt #3p       : Sku RateU BltRat [Blt/Ac]
DrpC "#3p"
RunqC "Select x.Sku,RateU,[Btl/Ac],CDbl(0) aS BtlRat into [#3p] from [SkuTaxBy3rdParty] x inner join Sku a on a.Sku=x.Sku"
RunqC "Update [#3p] set BtlRat =RateU/[Btl/Ac]"  ' RateU is in /Ac
End Sub
Private Sub WBH_HshRatBch86_OH(A As Ymd)
DrpC "#OHBch86"
Dim Q$: Q = "Select Distinct Sku,BchNo" & _
" into [#OHBch86]" & _
" from [OH]" & _
MH.TbOH.WhYmd(A) & " and Btl>0 and Co=86"
RunqC Q
End Sub
Private Sub WBH_HshRatBch86(A As Ymd)
'Oup : #RatBch86 = Sku BchNo PermitD Rate Permit PermitDate
DrpTTC "#RatBch86 #A"

RunqC "Select Distinct x.Sku,x.BchNo,Max(x.PermitD) as PermitD" & _
" Into [#A]" & _
" From (PermitD x" & _
" Inner Join [#OHBch86] a on a.BchNo=x.BchNo and a.Sku=x.Sku)" & _
" group by x.Sku,x.BchNo"

'Crt @BchRat: PermitD Sku BchNo Permit PermitDate
RunqC "Select x.Sku,x.BchNo,a.Rate,a.PermitD,a.Permit,PermitDate" & _
" Into [#RatBch86]" & _
" From ([#A] x" & _
" Inner Join [PermitD] a on a.PermitD=x.PermitD)" & _
" Inner Join [Permit]  b on a.Permit=b.Permit"
'Drp
DrpC "#A"
End Sub
Private Sub WBH_HshRatLas86(A As Ymd)
'Oup: Sku BchNo LasBchRate LasPermitD LasPermit LasPermitDate
'Inp: #OHBch86  : Sku BchNo
'Inp: #RatBch86 : Sku BchNo ...
DrpTTC "##Sku ##SkuPermitD"
'Crt [##Sku : Sku] These are Sku86-with-OH, but not may have some batch does not have batch-rate.
RunqC "Select Distinct x.Sku" & _
" into [##Sku]" & _
" from [#OHBch86] x" & _
" left join [#RatBch86] a on a.Sku=x.Sku and a.BchNo=x.BchNo" & _
" where a.Sku is null"

'Crt [##SkuPermitD] : Sku PermitD Only] those Sku86-with-OH with no Bch-rate
RunqC "Select Distinct x.Sku,Max(x.PermitD) as PermitD" & _
" Into [##SkuPermitD]" & _
" From (PermitD x" & _
" Inner Join [##Sku] a on a.Sku=x.Sku)" & _
" group by x.Sku"

'Crt #RatLas86: PermitD Sku BchNo Permit PermitDate
RunqC "Select" & _
" x.Sku                      ," & _
" a.BchNo                    ," & _
" Rate       as LasBchRat    ," & _
" x.PermitD  as LasPermitD   ," & _
" a.Permit   as LasPermit    ," & _
" PermitDate as LasPermitDate," & _
" Into [#RatLas86]" & _
" From ([##SkuPermitD] x" & _
" Inner Join [PermitD] a on a.PermitD=x.PermitD)" & _
" Inner Join [Permit]  b on a.Permit=b.Permit"

DrpTTC "##Sku ##SkuPermitD"
End Sub
Private Function WWSampYmd() As Ymd: WWSampYmd = MH.TbOH.YmdLas: End Function
Private Sub WO_RatPac(A As Ymd) ' Crt @PacRatD @PacRat
'Inp: Tb-SkuRepackMulti = SkuNew SkuFm FmSkuQty
'Inp: Tb-OH             = Sku Co Btl ..
'Oup: @PacRatD   = Sku Des FmSku FmDes FmQty FmSkuAcRat RefPermit RefBchNo
'Oup: @PacRat    = Sku Des PacRat FmSkuCnt (PacRat is /Ac)
DrpTTC "#WiOHParSku #ParChdSku #FmSkuRat #FmSkuPmi1 #FmSkuPmi2 @PacRatD @PacRat"
Dim Wh$: Wh = MH.TbOH.WhYmd(A)
'Crt #WiOHParSku : Sku
'    ===========
RunqC FmtQQ("Select Distinct Sku as NewSku" & _
" into [#WiOHParSku]" & _
" from [OH]" & _
" ? and Btl>0 and Co=86 and Sku in (Select Distinct SkuNew from SkuRepackMulti)", Wh)

'Crt #ParChdSku : Sku FmSku | FmQty
'    ==========
RunqC "Select SkuNew as NewSku, SkuFm as FmSku, FmSkuQty as FmQty" & _
" into [#ParChdSku]" & _
" from [SkuRepackMulti] x" & _
" where x.SkuNew in (Select NewSku from [#WiOHParSku])"

'Crt #FmSkuPmi1 : FmSku | Date_Id
'    ==========
RunqC "Select Sku as FmSku,Max(Format(PermitDate,'YYYY-MM-DD') & '_' & x.Permit) as Date_Id" & _
" Into [#FmSkuPmi1]" & _
" From PermitD x" & _
" Inner join Permit a on x.Permit=a.Permit" & _
" where Sku in (Select FmSku from [#ParChdSku])" & _
" Group By Sku"

'Crt #FmSkuPmi2 : FmSku | PermitDate Permit
'    ==========
RunqC "Select FmSku,CDate(Left(Date_Id,10)) as PermitDate, CLng(Mid(Date_Id,12)) as Permit" & _
" Into [#FmSkuPmi2]" & _
" From [#FmSkuPmi1]"

'Crt #FmSkuRat : FmSku | PermitDate Permit Rate
'    =========
RunqC "Select x.FmSku,x.Permit,x.PermitDate,Rate" & _
" into [#FmSkuRat]" & _
" from ([#FmSkuPmi2] x" & _
" inner join [PermitD] a on a.Sku=x.FmSku and a.Permit=x.Permit)"

'Crt @PacRatD : Sku FmSku Permit Rate
'    ========
RunqC "Select x.NewSku,'' As Des,x.FmSku,'' As FmSkuDes,x.FmQty,Rate,Permit,PermitDate" & _
" Into [@PacRatD]" & _
" From [#ParChdSku] x" & _
" left join [#FmSkuRat] a on x.FmSku = a.FmSku "
RunqC "Update [@PacRatD] x inner join Sku a on a.Sku=x.NewSku set x.Des=a.SkuDes"
RunqC "Update [@PacRatD] x inner join Sku a on a.Sku=x.FmSku set x.FmSkuDes=a.SkuDes"

'Crt @PacRat : SkuNew Rate
'    ========
RunqC "Select Distinct x.NewSku,'' as Des,Sum(FmQty*x.Rate) as Rate,Count(*) as FmSkuCnt" & _
" Into [@PacRat]" & _
" From [@PacRatD] x" & _
" Group by NewSku"
RunqC "Update [@PacRat] x inner join Sku a on a.Sku=x.NewSku set x.Des=a.SkuDes"

'-- Stp: Drp
DrpTTC "#WiOHParSku #ParChdSku #FmSkuRat #FmSkuPmi1 #FmSkuPmi2"
End Sub

Private Sub WO_RatBchMO(A As Ymd)
'Oup: [@BchRatMO]: Sku BchNo Val Btl | IsZHT0 [HKD/MOP] [MOP/Litre] [HKD/XXX] [XXX] [XXX/Btl]   with supporting....
'Fm   [OH]         Sku BchNo Val Btl       where Wh and Co=87
'Fm   [Sku]        Sku TaxRateMO          ! TaxRateMO>0
'## Stp    Oup       Stru
' 1 TmpOH  $OH       From [OH]  for those Wh and Co=87
' 2 TmpSku #Sku      From [Sku] for those TaxRateMO>0
' 3 TmpMO  @BchRatMO
'Tmp  [@BchRatMO]
DoCmd.SetWarnings False
StsQry "BchRatMO"

Dim Wh$: Wh = MH.TbOH.WhYmd(A)

'-- 1 Stp-TmpOH ===========================================================
RunqC "Select Sku,BchNo,Val,Btl into [$OH] from OH" & Wh & " and Co=87"
RunqC "Alter Table [$OH] add column " & _
"IsZHT0      yesno," & _
"[HKD/MOP]   Double," & _
"[MOP/Litre] Currency," & _
"[HKD/XXX]   Double," & _
"[XXX]       Text(3)," & _
"[XXX/Btl]   Currency "

RunqC "Update [$OH] x,MacauRatePrm a set " & _
"  IsZHT0     =False," & _
"x.[HKD/MOP]  =a.[HKD/MOP]," & _
"x.[MOP/Litre]=a.[MOP/Litre]," & _
"x.[HKD/XXX]  =a.[HKD/XXX]," & _
"x.[XXX]      =a.[XXX]"

'-- 2 Stp-TmpSku ========================================================
RunqC "Select Sku into [#Sku] from Sku where Nz(TaxRateMO,0)<>0"
RunqC "Update [$OH] x inner join [#Sku] a on a.Sku=x.Sku set x.IsZHT0=true"
RunqC "Update [$OH] x inner join MacauOverRideRate a on a.Sku=x.Sku and a.BchNo=x.BchNo set x.[XXX/Btl]=a.[XXX/Btl]"

'-- 3 Stp-OupBchRatMO ==================================================
RunqC "Select * into [@BchRatMO] from [$OH] where IsZHT0 or (not [XXX/Btl] is null)"
RunqC "Alter Table [@BchRatMO] add column [Btl/AC] Integer,[Litre/Btl] Double,Litre Double,LitreHKD Currency,[10%A] Currency,[10%B] Currency,[10%HKD] Currency,HKD Currency,[HKD/AC] Double,BchRat Double"
RunqC "Update [@BchRatMO] x inner join [Sku] a on a.Sku=x.Sku set x.[Litre/Btl]=a.[Litre/Btl],x.[Btl/AC]=a.[Btl/AC]"
'
RunqC "Update [@BchRatMO] set Litre    = Btl * [Litre/Btl]"
RunqC "Update [@BchRatMO] set LitreHKD = Litre * [MOP/Litre] / [HKD/MOP]"
RunqC "Update [@BchRatMO] set [10%A]   = [XXX/Btl] * Btl * 0.1"
RunqC "Update [@BchRatMO] set [10%B]   = [Val] * 0.1"
RunqC "Update [@BchRatMO] set [10%HKD] = IIF(Isnull([XXX/Btl]),[10%B],[10%A])"
RunqC "Update [@BchRatMO] Set HKD      = [LitreHKD] + [10%HKD]"
RunqC "Update [@BchRatMO] Set [HKD/AC] = HKD/Btl*[Btl/AC]"
RunqC "Update [@BchRatMO] SEt BchRat   = [HKD/AC]"
RunqC "Alter Table [@BchRatMO] add column SkuDes Text(50)"
RunqC "Update [@BchRatMO] x inner join Sku a on x.SKu=a.Sku set x.SkuDes = a.SkuDes"

'-- Rename [@BchRatMO] 4 fields with [Btl] into [Bott] due to template is using [Bott]
RenFldFf "@BchRatMO", _
    "Btl  XXX/Btl  Btl/AC  Litre/Btl", _
    "Bott XXX/Bott Bott/AC Size"

'== 4 Stp-DrpTmp
DrpTTC "#Sku"
End Sub

Private Sub WO_RatZHT0()
'Oup: @RatZHT0 (Co Sku  SkuDes   Uom       [Rate/Uom] [Btl/AC] [Unit/AC] ==> [Rate/AC] [RatZHT0]
'Fm:  Sku       x  x    SkuDes   TaxUomHK  TaxRateHK   x        x
'                       SkuDes   TaxUomMO  TaxRateMO                                           <== TaxUomHK is in HKD/Uom
'     Calc                                                                    xx     xx        <== They are same
'               RatZHT0 = Rate/Uom

'-- @RatZHT0: 7 Fields from Sku
DoCmd.SetWarnings False
DrpC "@RatZHT0"
Sts "RatZHT0"
Dim A$: A = "Insert into [@RatZHT0]" & _
                 " (Co,Sku,SkuDes,            Uom,             [Rate/Uom],[Btl/AC],[Unit/AC]) "
RunqC "Select 86 As Co,Sku,SkuDes,TaxUomHK as Uom,TaxRateHK As [Rate/Uom],[Btl/AC],[Unit/AC] into [@RatZHT0] From [SKU] where Nz(TaxRateHK,0)<>0"
RunqC A & _
             "Select 87 As Co,Sku,SkuDes,TaxUomMO as Uom,TaxRateMO As [Rate/Uom],[Btl/AC],[Unit/AC]                 From [SKU] where Nz(TaxRateMO,0)<>0"

'-- @RatZHT0: Add 2 fields [Rate/AC] [RatZHT0]
RunqC "Alter Table [@RatZHT0] add column [Rate/AC] Double, [RatZHT0] Double"

'-- @RatZHT0: Update [Rate/AC]
Select Case 1
Case 1
    'Not work:
    RunqC "Update [@RatZHT0] set [Rate/AC]=[Rate/Uom]*[Unit/AC] Where Uom='PCE'" 'PCE is Set
    RunqC "Update [@RatZHT0] set [Rate/AC]=[Rate/Uom]           Where Uom='CA'"  'CA  is AC
    RunqC "Update [@RatZHT0] set [Rate/AC]=[Rate/Uom]*[Btl/AC]  Where Uom='COL'" 'COL is Btl
Case 2
    With CurrentDb.OpenRecordset("Select [Rate/AC],[Rate/Uom],[Btl/AC],[Unit/AC],Uom From [@RatZHT0] where Uom in ('PCE','CA','COL')")
        While Not .EOF
            .Edit
            Select Case !UOM
            Case "PCE": .Fields("Rate/AC").Value = .Fields("Rate/Uom").Value * .Fields("Unit/AC").Value
            Case "CA":  .Fields("Rate/AC").Value = .Fields("Rate/Uom").Value
            Case "COL": .Fields("Rate/AC").Value = .Fields("Rate/Uom").Value * .Fields("Btl/AC").Value
            End Select
            .Update
            .MoveNext
        Wend
    End With
End Select
RunqC "Update [@RatZHT0] set [RatZHT0]=[Rate/AC]"
End Sub

Private Sub WR_GenFx(Fx$)
DltFfnIf Fx
ExpAttTpC ".Io.Tp.xlsx", Fx
Dim B As Workbook: Set B = WbFx(Fx)
RfhLoWbRsC B
Dim L As ListObject: Set L = LoFst(B.Sheets("Data"))
SetLoFmllny L, MH.MB52IO.FmllnyWsDta
Maxv B.Application
End Sub
