Attribute VB_Name = "MxDao_Fea_Lim_LnkImp_SpSamp"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Fea_LnkImp_SpSamp."

Function LnkImpSpSamp() As String()
Erase XX
X "*Spec :LnkImp MB52 *Inp- FbTbl- FxTbl- Tbl.Where- *Stru MusHasRecTbl"
X " Inp::(Inpn,Ffn)+"
X " FbTbl::(FbTbn,Stru..)*"
X " FxTbl::(FxTbn,?Inpnw,?Stru)*  "
X "         Inpnw is Inpnw is Inpn-dot-Wsn.  It is optional.  Inpn will use Fxin and Wsn will use sheet1"
X " Tbl.Where::(Inpn,Bepr)*                  The Bepr is using Extn in Sql-Bepr"
X " Stru::(Stru,(Intn,?Ty,?Extn))+           "
X "          Ty is (Dbl | Txt Dbl|Txt Dte)"
X "          Extn is a term, must quoated in []"
X " MustHasRec::(Inpn..|*AllInp)"
X "          *AllInpn all Inpn should have record"
X "Inp"
X " DutyPay C:\Users\User\Desktop\SAPAccessReports\DutyPrepay5\DutyPrepay5_Data.mdb"
X " ZHT0    C:\Users\user\Desktop\Mhd\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\Pricing report(ForUpload).xls"
X " MB52    C:\Users\user\Desktop\Mhd\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\2018\MB52 2018-01-30.xls"
X " Uom     C:\Users\user\Desktop\Mhd\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\sales text.xlsx"
X " GLBal   C:\Users\user\Desktop\Mhd\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\DutyPrepayGLTot.xlsx"
X "FbTbl"
X " --  Fbn T.."
X " DutyPay Permit PermitD"
X "FxTbl "
X " -- FxTbn Inpnw    Stru"
X " ZHT086  ZHT0.8600 ZHT0"
X " ZHT087  ZHT0.8700 ZHT0"
X " MB52"
X " Uom"
X " GLBal"
X "Tbl.Where"
X " MB52 Plant='8601' and [Storage Location] in ('0002','')"
X " Uom  Plant='8601'"
X "Stru Permit"
X " Permit"
X " PermitNo"
X " PermitDate"
X " PostDate"
X " NSku"
X " Qty"
X " Tot"
X " GLAc"
X " GLAcName"
X " BankCode"
X " ByUsr"
X " DteCrt"
X " DteUpd"
X "Stru PermitD"
X " PermitD"
X " Permit"
X " Sku"
X " SeqNo"
X " Qty"
X " BchNo"
X " Rate"
X " Amt"
X " DteCrt"
X " DteUpd"
X "Stru ZHT0"
X " Sku       Txt Material    "
X " CurRateAc Dbl [     Amount]"
X " VdtFm     Txt Valid From  "
X " VdtTo     Txt Valid to    "
X " HKD       Txt Unit        "
X " Per       Txt per         "
X " CA_Uom    Txt Uom         "
X "Stru MB52"
X " Sku    Txt Material          "
X " Whs    Txt Plant             "
X " Loc    Txt Storage Location  "
X " BchNo  Txt Batch             "
X " QInsp  Dbl In Quality Insp#  "
X " QUnRes Dbl UnRestricted      "
X " QBlk   Dbl Blocked           "
X " VInsp  Dbl Value in QualInsp#"
X " VUnRes Dbl Value Unrestricted"
X " VBlk   Dbl Value BlockedStock"
X " VBlk2  Dbl Value BlockedStock1"
X " VBlk1  Dbl Value BlockedStock2"
X "Stru Uom"
X " Sc_U    Txt SC "
X " Topaz   Txt Topaz Code "
X " ProdH   Txt Product hierarchy"
X " Sku     Txt Material            "
X " Des     Txt Material Description"
X " AC_U    Txt Unit per case       "
X " SkuUom  Txt Base Unit of Measure"
X " BusArea Txt Business Area       "
X "Stru GLBal"
X " BusArea Txt Business Area Code"
X " GLBal   Dbl                   "
X "Stru SkuTaxBy3rdParty"
X " SkuTaxBy3rdParty "
X "Stru SkuNoLongerTax"
X " SkuNoLongerTax"
X "MustHasRecTbl"
X " *AllInp"
LnkImpSpSamp = XX
End Function
