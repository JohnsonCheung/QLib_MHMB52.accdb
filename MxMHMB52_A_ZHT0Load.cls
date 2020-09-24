VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_ZHT0Load"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_ZHT0Load."
Const HKCusNo& = 5007964
Const MOCusNo& = 5007960
Private WWIo As New MxMHMB52_A_ZHT0Io
Sub LoadZHT0()
Dim Fxi$: Fxi = WWIo.Fxi
Dim Fxw$: Fxw = WWIo.Fxw
ChkFfnExi Fxi, CSub, "Forecast input file"
If Not Start("Start load ZHT0 tax rate") Then Exit Sub
WFxWrk Fxw
WImp Fxw
DrpC ">ZHT0"
End Sub

Private Sub WImp(Fxw$) 'Import Fxw into Sku..
' Fxw  -> [>ZHT0] -> #IZHT0 -> (#MO #HK) -> Sku
' #MO where Customer = MOCusNo
' #HK where Customer = HKCusNo

'#IZHT0 from >ZHT0
'#1 Crt #MO #HK             ::[Sku TaxRatePerUom SapUom] where
'#2 Crt #IZHT0 from #MO #HK
'#3 Upd Sku from #IZHT0
'#4   Prompt msg
'#5 Drp #MO #HK #IZHT0
Const T$ = "#IZHT0"

LnkFxwC Fxw, "Sheet1", ">ZHT0"
WTmpIZHT0
DoCmd.SetWarnings False

'-- Crt #HK #MO -> Crt #IZHT0
RunqC "select Material As SKU, CCur(Amount) As TaxRatePerUom, Uom as SapUom" & _
" into `#MO`" & _
" from [#IZHT0] where Customer=" & MOCusNo
RunqC "select Material As SKU, CCur(Amount) As TaxRatePerUom, Uom as SapUom" & _
" into `#HK`" & _
" from [#IZHT0] where Customer=" & HKCusNo

'-- Upd/Ins tb-Sku by [#IZHT0]
RunqC "Update Sku set TaxRateHK=Null,TaxRateMO=Null,TaxUOMHK=Null,TaxUOMMO=Null,DteRUpdTaxRate=Null"
Dim Tim$: Tim = Format(Now, "YYYY-MM-DD HH:MM:SS")
RunqC "Update Sku x inner join `#HK` a on x.Sku=a.Sku set x.TaxRateHK=a.TaxRatePerUom,x.TaxUOMHK=a.SapUom, DteRUpdTaxRate=#" & Tim & "#"
RunqC "Update Sku x inner join `#MO` a on x.Sku=a.Sku set x.TaxRateMO=a.TaxRatePerUom,x.TaxUOMMO=a.SapUom, DteRUpdTaxRate=#" & Tim & "#"
'-- Show message ========================================================================
Dim NHK%, NMO%
NMO = NRecTC("#MO")
NHK = NRecTC("#HK")
MsgBox "[" & NHK & "] HK Sku" & vbCrLf & "[" & NMO & "] Macau Sku have been imported." & vbCrLf & vbCrLf & "All old ZHT0 rate are totally replaced", vbInformation
DrpTTC "#HK #MO #IZHT0"
End Sub

Private Sub WTmpIZHT0() ' Crt [>ZHT0] from [#IZHT0].  Dlt rec for columns have invalid value: col[Customer Uom CnTy]..
'#1 DcDrs-Customer must be 5007960 5007964
'#2 DcDrs-CnTy     must be ZHT0
'#3 DcDrs-Uom      must in [COL PCE CA]
DoCmd.SetWarnings False
RunqC "Select * into [#IZHT0] from [>ZHT0]"

Const WhCnTy$ = "Nz(CnTy,'')<>'ZHT0'"
Const WhUom$ = "Not Uom in ('COL','PCE','CA')"
Const WhCus$ = "Not Customer in (5007960,5007964)"

Dim CntQ$: CntQ = SqlSelCnt("#IZHT0") & " where "
Dim NCnTy%: NCnTy = ValQC(CntQ & WhCnTy)
Dim NUom%:  NUom = ValQC(CntQ & WhUom)
Dim NCus%:  NCus = ValQC(CntQ & WhCus)

Select Case True
Case NCnTy <> 0, NUom <> 0, NCus <> 0
    MsgBox "There are records with invalid data in fields.  These records are ignored:" & vbCrLf & _
    "CnTy=[" & NCnTy & "]" & vbCrLf & _
    "Uom=[" & NCnTy & "]" & vbCrLf & _
    "Customer=[" & NCus & "]"
    Const DltQ$ = "Delete * from [#IZHT0] where "
    If NCnTy > 0 Then RunqC DltQ & WhCnTy
    If NUom > 0 Then RunqC DltQ & WhUom
    If NCus > 0 Then RunqC DltQ & WhCus
End Select
End Sub

Private Sub WFxWrk(FxWrk$) ' Crt @FxWrk from !WWIo.Fxi
DltFfnIf FxWrk
Dim Wb As Workbook: Set Wb = WbFx(WWIo.Fxi)
Dim Ws As Worksheet: Set Ws = WsFst(Wb)
'--Delete first 4 rows
RgWsRR(Ws, 6, 7).Delete
RgWsRR(Ws, 1, 4).Delete
CWs(Ws, "D").Delete
CWs(Ws, "A").Delete
CWs(Ws, "L").Delete
RgWsRC(Ws, 1, "F").Value = Trim(RgWsRC(Ws, 1, "F").Value)
CWs(Ws, "D").NumberFormat = "@"                 'format Material as text
'To get ride of the blank columns & lines when using LasCells in LoRg
Dim Lo As ListObject: Set Lo = LoWsDta(Ws)
Dim R%(): R = WRny(Lo)
Dim J%: For J = UB(R) To 0 Step -1
    Lo.ListRows(R(J)).Range.Delete
Next
Wb.SaveAs FxWrk, XlFileFormat.xlOpenXMLWorkbook
QuitWb Wb
End Sub

Private Function WRny(Lo As ListObject) As Integer() ' ret @Rny::Row-no-Ay of @Lo, if col-Customer2 has EMpty value.
Dim Sq(): Sq = Lo.ListColumns("Customer2").DataBodyRange.Value
Dim R%: For R = 1 To UBound(Sq, 1)
    Dim V: V = Sq(R, 1)
    Select Case True
    Case IsEmpty(V), V = "Customer"
        PushI WRny, R
    End Select
Next
End Function
