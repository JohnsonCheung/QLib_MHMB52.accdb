VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_MB52Tp_Fmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Lc_Wdt(Lo, "Val", 12
'FmtLcWdt Lo, "Btl AC SC", 8
'Lc_Wdt(Lo, "Co", 4
'FmtLcWdt Lo, "YY MM DD", 2
'FmtLcWdt Lo, "Stream", 6
'FmtLcWdt Lo, "PHBus", 2
'FmtLcWdt Lo, "BusArea", 6
'
'FmtLcWdt Lo, "TaxLoc TaxItm", 3
'
'Lc_Wdt(Lo, "NmYpStk", 20
'FmtLcWdt Lo, "YpStk", 2
'Lc_Wdt(Lo, "Sku", 8
'FmtLcWdt Lo, "Litre/Btl StkUnit", 7
'FmtLcWdt Lo, "Btl/AC Unit/SC Unit/AC Btl/Unit", 5
'FmtLcWdt Lo, "ml/Btl", 7
'FmtLcWdt Lo, "Btl/AC' Litre/SC Btl Btl/AC Unit/SC Unit/AC Btl/Unit", 5
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_MB52Tp_Fmt."

Sub Fmt()
Dim X As Excel.Application: Set X = XlsNw
Dim B As Workbook: Set B = X.Workbooks.Open(MH.MB52Tp.Tp)
Maxv B.Application
WFmtWsDta B.Sheets("Data")
WFmtPt B
'SetWbOutLinSum B
'SetWbNoAutoColWdt B
B.Save
Maxv B.Application
Done
End Sub
Private Sub B_MB52Tp_Fmt()
Dim X As Excel.Application: Set X = XlsNw
Dim B As Workbook: Set B = X.Workbooks.Open(MH.MB52Tp.Tp)
Dim S As Worksheet, Pt As PivotTable
Maxv B.Application
Stop
Dim Wsn: For Each Wsn In Tmy("Bch Sku Naming Brand BusArea Quality [Quality Group]")
    Set S = B.Sheets(Wsn)
    If S.PivotTables.Count >= 1 Then
        Set Pt = FstPt(S)
        FmtPivDfFmt Pt, "[Sum of BchAmt] [Sum of ZHT0Amt]", "$#,##0;-$#,##0;#"
        FmtPivDfFmt Pt, "[Average of RatDif] [Sum of AmtDif]", "$#,##0;-$#,##0;#"
    End If
Next
B.Save
Maxv B.Application
Done
End Sub
Private Sub WFmtWsDta(Ws As Worksheet)
Ws.Activate
Dim Lo As ListObject: Set Lo = LoFst(Ws)
WFmtFmt Lo
'WFmtTot Lo
'WFmtWdt Lo
'WFmtFml Lo
'WFmtLvl Lo
End Sub
Private Sub WFmtFml(Lo As ListObject): SetLoFmllny Lo, MH.MB52IO.FmllnyWsDta: End Sub
Private Sub WFmtFmt(Lo As ListObject)
'== Ws-Data: Format Column ===
SetLcFmt Lo, "Litre/Btl", "0.000;-0.000;#"
SetLcFmt Lo, "Btl/AC", "#"
SetLcFmt Lo, "Unit/AC", "#"
SetLcFmt Lo, "Unit/SC", "0.00"
SetLcFmt Lo, "Btl/Unit", "#"

'--Siz
SetLcFmt Lo, "ml/Btl", "#,###"
SetLcFmt Lo, "Btl/AC'", "#"
SetLcFmt Lo, "Litre/SC", "#,##0.0"
'
SetLcFmt Lo, "Val", "$#,###"
SetLcFmt Lo, "Btl", "#,###"
SetLcFmt Lo, "SC", "#,##0.0"
SetLcFmt Lo, "AC", "#.##0.0"
SetLcFmt Lo, "ScUPr", "$#,###"
SetLcFmt Lo, "AcUPr", "$#,###"
SetLcFmt Lo, "BtlUPr", "$#,###"
SetLcFmt Lo, "BchRat", "$#,###;-$#,###;"""""
SetLcFmt Lo, "RatZHT0", "$#,###;-$#,###;"""""
SetLcFmt Lo, "BchAmt", "$#,###;-$#,###;"""""
SetLcFmt Lo, "ZHT0Amt", "$#,###;-$#,###;"""""

SetLcFmt Lo, "RatDif", "$#,###;-$#,###;"""""
SetLcFmt Lo, "AmtDif", "$#,###;-$#,###;"""""
End Sub
Private Sub WFmtTot(Lo As ListObject)
SetLccAsSum Lo, "Val Btl AC SC"
SetLccAsAvg Lo, "ScUPr BtlUPr AcUPr"
SetLccAsAvg Lo, "BchRat RatZHT0"
SetLccAsSum Lo, "BchAmt ZHT0Amt"
SetLccAsSum Lo, "AmtDif"
SetLcAsCnt Lo, "SkuDes"
End Sub
Private Sub WSetWdt(L As ListObject)
SetLcWdt L, "12 Val"
SetLcWdt L, "8 Btl AC SC"
SetLcWdt L, "4 Co"
SetLcWdt L, "2 YY MM DD"
SetLcWdt L, "6 Stream"
SetLcWdt L, "2 PHBus"
SetLcWdt L, "6 BusArea"
SetLcWdt L, "3 TaxLc TaxItm"
SetLcWdt L, "20 NmYpStk"
SetLcWdt L, "2 YpStk"
SetLcWdt L, "50 SkuDes"
SetLcWdt L, "8 Sku"
SetLcWdt L, "7 Litre/Btl StkUnit"
SetLcWdt L, "5 Btl/AC Unit/SC Unit/AC Btl/Unit"
SetLcWdt L, "7 ml/Btl"
SetLcWdt L, "5 Btl/AC' Litre/SC Btl Btl/AC Unit/SC Unit/AC Btl/Unit"
End Sub
Private Sub WFmtLvl(L As ListObject)
SetLcLvl L, "YY MM DD"
SetLcLvl L, "PHBus BusArea NmYpStk YpStk"
SetLcLvl L, "Sku Litre/Btl Btl/AC  Unit/SC Unit/AC StkUnit"
SetLcLvl L, "BchAmt  ZHT0Amt RatDif  AmtDif  BchNo   BchRatTy"
SetLcLvl L, "PH  PHNam   PHBrd   PHQGp   PHQly   PHSStm  PHSBus  PHSrt1  PHSrt2  PHSrt3  PHSrt4  PHL1    PHL2    PHL3    PHL4"
End Sub

Private Sub WFmtPt(Wb As Workbook)
WFmtWsBchPt Wb.Sheets("Bch1")
Dim PhItm: For Each PhItm In MH.PH.Itmy7
    WFmtPtPHItm PhItm
Next
End Sub
Private Sub B_WFmtWsBchPt()
Dim X As New Excel.Application
Maxv X
Dim Wb As Workbook: Set Wb = X.Workbooks.Open(MH.MB52Tp.Tp)
Dim Ws As Worksheet: Set Ws = Wb.Sheets("Bch1")
WFmtWsBchPt Ws
Wb.Save
Done
End Sub
Private Sub WFmtWsBchPt(Ws As Worksheet)
Dim Pt As PivotTable: Set Pt = FstPt(Ws)
WFmtWsBchPtWdt Pt
WFmtBchWsPtFmt Pt
End Sub

Private Sub WFmtWsBchPtWdt(Pt As PivotTable)
Const W_PH4Des% = 20
Const W_Sku% = 8
Const W_SkuDes% = 50
Const W_cv_mlBtl% = 6
Const W_cv_BtlAc% = 4
Const W_cv_LitreSc% = 5
Const W_df_Val% = 8
Const W_df_Val2% = 12
FmtPivRfWdt Pt, "PHNam PHBrd PHQGp PHQly", W_PH4Des
FmtPivRfWdt Pt, "Sku", W_Sku
FmtPivRfWdt Pt, "SkuDes", W_SkuDes
FmtPivRfWdt Pt, "ml/Btl", W_cv_mlBtl
FmtPivRfWdt Pt, "Btl/AC", W_cv_BtlAc
FmtPivRfWdt Pt, "Litre/SC", W_cv_LitreSc
FmtPivDfWdt Pt, "[Sum of Val] [Sum of AC] [Sum of SC] [Sum of Btl] [Average of BtlUPr] [Average of AcUPr] [Average of ScUPr]" & _
" [Average of BchRat] [Average of RatDif] [Average of RatZHT0]", W_df_Val
FmtPivDfWdt Pt, "[Sum of BchAmt] [Sum of ZHT0Amt] [Sum of AmtDif]", W_df_Val2
End Sub
Private Sub WFmtBchWsPtFmt(Pt As PivotTable)
'D PtDtaFny(Pt)

'4: Val+3Q
'3: Price
'2: Rate
'2: Amt
'1: RatDif
'1: AmtDif
FmtPivDfFmt Pt, "[Sum Of Val]", "$#,##0"
FmtPivDfFmt Pt, "[Sum of AC]", "#.##0"
FmtPivDfFmt Pt, "[Sum of SC]", "#,##0.0"
FmtPivDfFmt Pt, "[Sum of Btl]", "#,##0"

FmtPivDfFmt Pt, "[Average of AcUPr] [Average of BtlUPr] [Average of ScUPr]", "$#,##0"

FmtPivDfFmt Pt, "[Average of BchRat] [Average of RatZHT0]", "$#,##0"
FmtPivDfFmt Pt, "[Sum of BchAmt] [Sum of ZHT0Amt]", "$#,##0"

FmtPivDfFmt Pt, "[Average of RatDif] [Sum of AmtDif]", "$#,##0"
End Sub
Private Sub WCrtPt(WsDta As Worksheet)
'
'Const AtAdr$ = "Bch1!A5"
'Dim Pt As PivotTable
'Dim Pc As PivotCache
'ClrPt BchWb
'    set Pt =    WbWs(BchWs).PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
'        "Data", Version:=6).CreatePivotTable TableDestination:=AtAdr, TableName:="PivotTable3", _
'        DefaultVersion:=6
'    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Stream")
'        .Orientation = xlPageField
'        .Position = 1
'    End With
'    With ActiveSheet.PivotTables("PivotTable3").PivotFields("PHBus")
'        .Orientation = xlRowField
'        .Position = 1
'    End With
'    With ActiveSheet.PivotTables("PivotTable3").PivotFields("BusArea")
'        .Orientation = xlRowField
'        .Position = 2
'    End With
'    With ActiveSheet.PivotTables("PivotTable3").PivotFields("NmYpStk")
'        .Orientation = xlColumnField
'        .Position = 1
'    End With
'    ActiveSheet.PivotTables("PivotTable3").AddDataField ActiveSheet.PivotTables( _
'        "PivotTable3").PivotFields("Val"), "Sum of Val", xlSum
'    With ActiveSheet.PivotTables("PivotTable3")
'        .InGridDropZones = True
'        .RowAxisLayout xlTabularRow
'    End With
'    Range("A8").Select
'    ActiveSheet.PivotTables("PivotTable3").PivotFields("PHBus").Subtotals = Array( _
'        False, False, False, False, False, False, False, False, False, False, False, False)
'End Sub
'
End Sub

Private Sub WFmtPtPHItm(PhItm)
Dim Wsn$
Stop
End Sub

Private Function WWbTp() As Workbook: Set WWbTp = XlsNw.Workbooks.Open(MH.MB52Tp.Tp): End Function
