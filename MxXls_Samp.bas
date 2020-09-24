Attribute VB_Name = "MxXls_Samp"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_SMP."
Function samp_mhmb52rpt_Fx$():                        samp_mhmb52rpt_Fx = sampFfn("Samp Rpt MB52.xlsx"):      End Function
Function samp_mhmb52rpt_Wb() As Workbook:         Set samp_mhmb52rpt_Wb = WbFx(samp_mhmb52rpt_Fx):            End Function
Function samp_mhmb52pgm_Fb$():                        samp_mhmb52pgm_Fb = sampFfn("Samp Rpt MB52 Oup.accdb"): End Function
Function samp_mhmb52pgm_Db() As Database:         Set samp_mhmb52pgm_Db = Db(samp_mhmb52pgm_Fb):              End Function
Function samp_mhmb52rptdta_Pt() As PivotTable: Set samp_mhmb52rptdta_Pt = PtRg(samp_mhmb52rptdta_Rg):         End Function
Function samp_mhmb52rptdta_Ws() As Worksheet:  Set samp_mhmb52rptdta_Ws = samp_mhmb52rpt_Wb.Sheets("Data"):   End Function
Function samp_mhmb52rptdta_Lo() As ListObject: Set samp_mhmb52rptdta_Lo = LoFst(samp_mhmb52rptdta_Ws):        End Function
Function samp_mhmb52rptdta_Rg() As Range:      Set samp_mhmb52rptdta_Rg = samp_mhmb52rptdta_Lo.DataBodyRange: End Function
Function samp_mhmb52rpt_Lof() As String()
Dim O$()
PushI O, "Lo  Nm     *Nm"
PushI O, "Lo  Fld    *Fld.."
PushI O, "Ali Left   *Fld.."
PushI O, "Ali Right  *Fld.."
PushI O, "Ali Center *Fld.."
PushI O, "Bdr Left   *Fld.."
PushI O, "Bdr Right  *Fld.."
PushI O, "Bdr DcDrs    *Fld.."
PushI O, "Tot Sum    *Fld.."
PushI O, "Tot Avg    *Fld.."
PushI O, "Tot Cnt    *Fld.."
PushI O, "Fmt *Fmt   *Fld.."
PushI O, "Wdt *Wdt   *Fld.."
PushI O, "Lvl *Lvl   *Fld.."
PushI O, "Cor *Cor   *Fld.."
PushI O, "Fml *Fld   *Formula"
PushI O, "Bet *Fld   *Fld1 *Fld2"
PushI O, "Tit *Fld   *Tit"
PushI O, "Lbl *Fld   *Lbl"
samp_mhmb52rpt_Lof = O
End Function
