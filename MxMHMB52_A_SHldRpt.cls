VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_SHldRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_SHldRpt."
Const FF10Val$ = "ScCsg HkdCsg ScDf HkdDf ScDp HkdDp ScGit HkdGit ScTot HkdTot"
Const FF4Kpi$ = "StkDays StkMths RemSC TarStkMths"
Const FF14Var$ = "ScCsg HkdCsg ScDf HkdDf ScDp HkdDp F1 ScGit HkdGit F2 ScTot HkdTot"
Private Fun As New MxMHMB52_MHFun
Private MHNw As New MxMHMB52_MHNw
Private MFxo86$, MFxo87$, MFxyO2$()
Private MHasFcMH As Boolean, MHasFcUD As Boolean
Private WhOHYmd$
Private MXls As Excel.Application
Private MYpStkGit As Byte
Private MYmd As Ymd, Myy As Byte, Mmm As Byte, MYm As Ym
Private MWhOHYmd$

Private Sub B_W_GenFxo()
WIniSampXls
W_GenFxo 86
W_GenFxo 87
MXls.WindowState = xlMaximized
MXls.Windows.Arrange xlArrangeStyleVertical
Done
End Sub
Friend Sub B__Tst()
B_WFB_FmtA1
B_WOFc
B_WOStkDays
B_WOStkHld
B_W_GenFxo
B_WOStkHld_Add3_PhRollUpCol
B_W_Rpt
End Sub
Private Function WYmdSamp() As Ymd: WYmdSamp = MH.TbOH.YmdLas:   End Function
Private Sub WIniSamp():                        WIni WYmdSamp:    End Sub
Private Sub WIniSampXls():                     WIniXls WYmdSamp: End Sub
Private Sub B_WOStkHld_Add3_PhRollUpCol()
DoCmd.Close acTable, "#A"
RunqC "Select Sku,SkuDes into [#A] from Sku"
WOStkHld_Add3_PhRollUpCol "#A"
DoCmd.OpenTable "#A"
End Sub
Private Sub B_WFB_FmtA1()
WIniSampXls
WFB_FmtA1 WWbCo(86), 86
WMaxv
End Sub
Private Function WWbCo(Co As Byte) As Workbook
Set WWbCo = WbFxX(WFxoCo(Co), MXls)
End Function

Private Sub WMaxv(): Maxv MXls: End Sub
Private Sub B_W_Rpt()
WIniSampXls
MH.PhTmp5.Crt
W_Rpt 86
Maxv MXls
End Sub
Private Sub B_WOStkHld():  WIniSamp: WOStkHld 86:  End Sub
Private Sub B_WOStkDays(): WIniSamp: WOStkDays 86: End Sub

Private Function WAskOpnFxyO2() As Boolean: WAskOpnFxyO2 = AskOpnFxy(MFxyO2): End Function
Private Sub WIniXls(Ymd As Ymd):                           WIni Ymd, Xls:     End Sub
Private Sub WIni(Ymd As Ymd, Optional X As Excel.Application)
MYmd = Ymd
Mmm = Ymd.M
Myy = Ymd.Y
MYm = YmYmd(MYmd)
MFxo86 = MH.SHldIo.Fxo86(Ymd)
MFxo87 = MH.SHldIo.Fxo87(Ymd)
ReDim MFxyO2(1)
MFxyO2(0) = MFxo86
MFxyO2(1) = MFxo87
MHasFcUD = WHasFc("U")
MHasFcMH = WHasFc("M")
MWhOHYmd = MH.TbOH.WhYmd(MYmd)
MYpStkGit = MH.GitIO.YpStk
Set MXls = X
End Sub
Private Function WHasFc(Stm$) As Boolean
Dim Wh$: Wh = MH.FcWh.WhStmYm(MHNw.StmYm(Stm, Myy, Mmm))
WHasFc = HasRecQC("Select Top 1 VerYY  from FcSku" & Wh)
End Function
Sub RptSHld(Ymd As Ymd)
WIniXls Ymd
If WAskOpnFxyO2 Then Exit Sub
If Not MH.IsLoad.MB52(Ymd) Then Exit Sub
If Not MH.IsLoad.Git(Ymd) Then Exit Sub
If Not WIsNoFcAsk Then Exit Sub
DltFfnyIf MFxyO2
MH.SHldStkDaysCalc.Calc Ymd  ' Cannot put in Load_MB52, because, it is required to calculate the stock days to include the Git
MH.PhTmp5.Crt
Set MXls = New Excel.Application
W_Rpt 86
W_Rpt 87
OpnFxyX MFxyO2, MXls
End Sub
Private Function WIsNoFcAsk() As Boolean
If MHasFcMH And MHasFcUD Then Exit Function
Dim Msg$
    Dim O$()
    If Not MHasFcMH Then PushS O, "No MH forecast"
    If Not MHasFcUD Then PushS O, "No UD forecast"
    Msg = JnCrLf(O) & vbCrLf & vbCrLf & "[Ok]=Continue to generate report" & vbCrLf & "or [Cancel]"
WIsNoFcAsk = MsgBox(Msg, vbQuestion + vbOKCancel) = vbCancel
End Function
Private Sub WUpdTbReport()
Dim Wh$: Wh = MWhOHYmd
Select Case 2
Case 1: RunqC "Update Report x set DteGen=Now()" & Wh
Case 2
    With RsQC("Select DteGen from Report" & Wh)
        .Edit
        !DteGen = Now()
        .Update
    End With
End Select
End Sub
Private Sub W_Rpt(Co As Byte)
WOStkHld 86    '@StkHld{7}
WOStkDays 86  '@StkDays{7}
WOFc Co        '@Fc{7}
W_GenFxo Co
DrpTTC "#2Tot #A #B #Piv >MB52 #YpSku #StkHld $PH "
DrpPfxxC "@Fc @StkHld @StkDays"
DoCmd.SetWarnings True
End Sub

Private Sub WOStkHld(Co As Byte)
'Aim: Create @StkHld{7} from Tbl-OH
'10+4Val = ScCsg HkdCsg ScDf HkdDf ScDp  HkdDp F1 ScGit HkdGit F2 ScTot HkdTot StkDays StkMth RemSC StkMthTar
'                    #StkHld{7}    CpyAddPHAtr
'                    {7Key}        {7Atr}
'Oup: @StkHldStm   = Stm         | Stm
'Oup: @StkHldBus   = Stm BusArea | Stm PHSBus BusArea PHBus
'Oup: @StkHldL1    = Stm PHL1    | Stm Srt1 PHL1 PHNam
'Oup: @StkHldL2    = Stm PHL2    | Stm Srt2 PHL2 PHNam PHBrd
'Oup: @StkHldL3    = Stm PHL3    | Stm Srt3 PHL3 PHNam PHBrd PHQGp
'Oup: @StkHldL4    = Stm PHL4    | Stm Srt4 PHL4 PHNam PHBrd PHQGp PHQly
'Oup: @StkHldSku   = Sku         | Stm Srt4 PHL4 PHNam PHBrd PHQGp PHQly Sku SkuDes

'     Stp       Oup
'     Beg       OH       = YY MM DD Co YpStk SKu BchNo | Btl Val
'  1   Sum       #YpSku   = YpStk Sku | Sc V                From OH
'  2   Piv       #Piv        = Sku   {8SC/Hkd}            Sku is not unique
'  3   2Tot      #2Tot       = Sku    |  {8SC/Hkd}  {2Tot}   Sku is unique
'  4   Rollup    #StkHld{7}  = {7Key} | {10Val}
'  5   Add4Kpi   @StkHld{7}  = Sku | {10Val} {4Days}
'  6   CpyAddPHAtr @StkHld{7}  = {7Atr} | {10Val}
'  7   AddF1F2   @StkHld{7}
'  8   ReSeq
'  9   DrpTmp
'Ref-For-AddAtr: $PH{5} PHLBus PHLStm = Sku | Btl/Sc ..
DrpPfxxC "@StkHld"
DoCmd.SetWarnings False
'== 1 Stp-Sum OH & Git

'PX "AC      =[@Btl] / [@[Btl/AC]]"
'PX "SC      =[@AC] * [@[Unit/AC]] / [@[Unit/SC]]"
'From OH
Dim Wh$: Wh = MWhOHYmd
Dim YpStkGit%: YpStkGit = MH.GitIO.Pthi
RunqC "SELECT YpStk,x.Sku,Sum([Btl]/[Btl/AC]*[Unit/AC]/[Unit/SC]) AS SC, Sum(Val/1000) AS V" & _
" INTO [#YpSku]" & _
" FROM (OH x" & _
" LEFT JOIN [Sku] a ON x.Sku = a.Sku)" & Wh & _
" GROUP BY YpStk,x.Sku;"

'From GitDet
RunqC "Insert into [#YpSku] (YpStk,Sku,SC,V)" & _
" SELECT " & MH.GitIO.YpStk & " As YpStk ,x.Sku,Sum([Btl]/[Btl/AC]*[Unit/AC]/[Unit/SC]) AS SC, Sum(HKD/1000) AS V" & _
" FROM (GitDet x" & _
" LEFT JOIN [Sku] a ON x.Sku = a.Sku)" & Wh & _
" GROUP BY x.Sku;"

If False Then
    Debug.Print Co, "SC/V",
    With CurrentDb.OpenRecordset("Select Sum(x.SC) as SC,Sum(x.V) as V from [#YpSku] x ")
        Debug.Print .Fields(0).Value; .Fields(1).Value
    End With
    Stop
End If

'-- 2 Stp-Piv =====================================================================================
DrpC "#Piv"
RunqC "Create Table [#Piv] (Sku Text(20)," & _
" ScCsg double, HkdCsg currency," & _
" ScDF  double, HkdDF  currency," & _
" ScDP  double, HkdDp  currency," & _
" ScGIT double, HkdGIT currency)"

RunqC "INSERT INTO [#Piv] (Sku, HKDCsg, ScCsg) SELECT Sku, Sum(V), Sum(SC) FROM [#YpSku] Where YpStk IN (SELECT YpStk FROM YpStk Where YpCls='*Consignment') group by Sku"
RunqC "INSERT INTO [#Piv] (Sku, HKDDF , ScDF ) SELECT Sku, Sum(V), Sum(SC) FROM [#YpSku] Where YpStk IN (SELECT YpStk FROM YpStk Where YpCls='*DutyFree')    group by Sku;"
RunqC "INSERT INTO [#Piv] (Sku, HKDDP , ScDP ) SELECT Sku, Sum(V), Sum(SC) FROM [#YpSku] Where YpStk IN (SELECT YpStk FROM YpStk Where YpCls='*DutyPaid')    group by Sku;"
RunqC "INSERT INTO [#Piv] (Sku, HKDGit , ScGit) SELECT Sku, Sum(V), Sum(SC) FROM [#YpSku] Where YpStk IN (SELECT YpStk FROM YpStk Where YpCls='*Git')    group by Sku;"

'-- 3 Stp-2Tot................................................................
RunqC "SELECT Distinct x.Sku, " & _
" Sum(x.ScCsg) AS ScCsg, Sum(x.HKDCsg) AS HKDCsg," & _
" Sum(x.ScDF)  AS ScDF , Sum(x.HKDDF)  AS HKDDF  ," & _
" Sum(x.ScDP)  AS ScDP , Sum(x.HKDDP)  AS HKDDP  ," & _
" Sum(x.ScGIT) AS ScGIT, Sum(x.HKDGIT) AS HKDGIT" & _
" Into [#2Tot]" & _
" From [#Piv] x" & _
" Group by x.Sku"

RunqC "Alter Table [#2Tot] add column ScTot Double, HKDTot Currency"
RunqC "Update [#2Tot] set" & _
" ScTot=Nz(ScCsg,0)+Nz(ScDp,0)+Nz(ScDf,0)+Nz(ScGit,0)," & _
" HkdTot=Nz(HkdCsg,0)+Nz(HkdDp,0)+Nz(HkdDf,0)+Nz(HkdGit,0)"
If False Then
    With CurrentDb.OpenRecordset("Select Sum(ScTot),Sum(HkdTot) from [#2Tot]")
        Debug.Print "SC"; .Fields(0).Value
        Debug.Print "V "; .Fields(1).Value
    End With
    Stop
End If
'4 Stp-Rollup  ================================================================
'   Oup:#StkHld{7}
Const Sum$ = "Sum(x.ScCsg) as ScCsg,Sum(x.HkdCsg) as HkdCsg," & _
"Sum(x.ScDf) as ScDf,Sum(x.HkdDf) as HkdDf," & _
"Sum(x.ScDp) as ScDp,Sum(x.HkdDp) as HkdDp," & _
"Sum(x.ScGit) as ScGit,Sum(x.HkdGit) as HkdGit," & _
"Sum(x.ScTot) as ScTot,Sum(x.HkdTot) as HkdTot"

WOStkHld_Add3_PhRollUpCol "#2Tot"
RunqQQC "SELECT x.Sku,      ? INTO [#StkHldSku] FROM [#2Tot] x Group by Sku", Sum
RunqQQC "SELECT Stm,BusArea,? INTO [#StkHldBus] FROM [#2Tot] x Group by Stm,BusArea", Sum
RunqQQC "SELECT Stm,PHL4,   ? INTO [#StkHldL4]  FROM [#2Tot] x Group by Stm,PHL4", Sum
RunqQQC "SELECT Stm,Left(PHL4,7) as PHL3,? INTO [#StkHldL3] FROM [#StkHldL4] x Group by Stm,Left(PHL4,7)", Sum
RunqQQC "SELECT Stm,Left(PHL3,4) as PHL2,? INTO [#StkHldL2] FROM [#StkHldL3] x Group by Stm,Left(PHL3,4)", Sum
RunqQQC "SELECT Stm,Left(PHL2,2) as PHL1,? INTO [#StkHldL1] FROM [#StkHldL2] x Group by Stm,Left(PHL2,2)", Sum
RunqQQC "SELECT Stm,                     ? INTO [#StkHldStm] FROM [#StkHldL1] x Group by Stm", Sum
If False Then
    Stop
End If
'--5 Stp-Add4Kpi : StkDays/RemSC/StkMths/TarStkMths
'   ref: PHTarMths{7}
'   ref: PHStkDays{7}
Const QpAddDc$ = " add column StkDays integer, RemSC Double, StkMths Single, TarStkMths Single"
RunqC "Alter Table [#StkHldSku]" & QpAddDc
RunqC "Alter Table [#StkHldL4] " & QpAddDc
RunqC "Alter Table [#StkHldL3] " & QpAddDc
RunqC "Alter Table [#StkHldL2] " & QpAddDc
RunqC "Alter Table [#StkHldL1] " & QpAddDc
RunqC "Alter Table [#StkHldBus]" & QpAddDc
RunqC "Alter Table [#StkHldStm]" & QpAddDc
Dim WhYmd$: WhYmd = MWhOHYmd
Dim WhCo$: WhCo = WhFeq("Co", Co)
'-- Sku ----------------------------------------
RunqC "Select Sku,StkDays,RemSC into [#A] from PHStkDaysSku" & WhYmd
RunqC "Select Sku,TarStkMths    into [#B] from PHTarMthsSKU" & WhCo
RunqC "Update [#StkHldSku] x inner join [#A] a on a.Sku=x.Sku set x.StkDays=a.StkDays,x.RemSC=a.RemSC,x.StkMths=a.StkDays/12"
RunqC "Update [#StkHldSku] x inner join [#B] a on x.Sku=a.Sku set x.TarStkMths=a.TarStkMths"

'-- Bus ----------------------------------------
RunqC "Select Stm,BusArea,StkDays,RemSC into [#A] from PHStkDaysBus" & WhYmd
RunqC "Select Stm,BusArea,TarStkMths    into [#B] from PHTarMthsBus" & WhCo
RunqC "Update [#StkHldBus] x inner join [#A] a on x.BusArea=a.BusArea and x.Stm=a.Stm set x.StkDays=a.StkDays,x.RemSC=a.RemSC,x.StkMths=a.StkDays/12"
RunqC "Update [#StkHldBus] x inner join [#B] a on x.BusArea=a.BusArea and x.Stm=a.Stm set x.TarStkMths=a.TarStkMths"

'-- PHL4 ----------------------------------------
RunqC "Select Stm,PHL4,StkDays,RemSC into [#A] from PHStkDaysL4" & WhYmd
RunqC "Select Stm,PHL4,TarStkMths    into [#B] from PHTarMthsL4" & WhCo
RunqC "Update [#StkHldL4] x inner join [#A] a on x.PHL4=a.PHL4 and x.Stm=a.Stm  set x.StkDays=a.StkDays,x.RemSC=a.RemSC,x.StkMths=a.StkDays/12"
RunqC "Update [#StkHldL4] x inner join [#B] a on x.PHL4=a.PHL4 and x.Stm=a.Stm  set x.TarStkMths=a.TarStkMths"

'-- PHL3 ----------------------------------------
RunqC "Select Stm,PHL3,StkDays,RemSC into [#A] from PHStkDaysL3" & WhYmd
RunqC "Select Stm,PHL3,TarStkMths    into [#B] from PHTarMthsL3" & WhCo
RunqC "Update [#StkHldL3] x inner join [#A] a on a.PHL3=x.PHL3 and x.Stm=a.Stm  set x.StkDays=a.StkDays,x.RemSC=a.RemSC,x.StkMths=a.StkDays/12"
RunqC "Update [#StkHldL3] x inner join [#B] a on x.PHL3=a.PHL3 and x.Stm=a.Stm  set x.TarStkMths=a.TarStkMths"

'-- PHL2 ----------------------------------------
RunqC "Select Stm,PHL2,StkDays,RemSC into [#A] from PHStkDaysL2" & WhYmd
RunqC "Select Stm,PHL2,TarStkMths    into [#B] from PHTarMthsL2" & WhCo
RunqC "Update [#StkHldL2] x inner join [#A] a on a.PHL2=x.PHL2 and x.Stm=a.Stm  set x.StkDays=a.StkDays,x.RemSC=a.RemSC,x.StkMths=a.StkDays/12"
RunqC "Update [#StkHldL2] x inner join [#B] a on x.PHL2=a.PHL2 and x.Stm=a.Stm  set x.TarStkMths=a.TarStkMths"

'-- PHL1
RunqC "Select Stm,PHL1,StkDays,RemSC into [#A] from PHStkDaysL1" & WhYmd
RunqC "Select Stm,PHL1,TarStkMths    into [#B] from PHTarMthsL1" & WhCo
RunqC "Update [#StkHldL1] x inner join [#A] a on a.PHL1=x.PHL1 and x.Stm=a.Stm  set x.StkDays=a.StkDays,x.RemSC=a.RemSC,x.StkMths=a.StkDays/12"
RunqC "Update [#StkHldL1] x inner join [#B] a on x.PHL1=a.PHL1 and x.Stm=a.Stm  set x.TarStkMths=a.TarStkMths"

'-- Stm
RunqC "Select Stm,StkDays,RemSC into [#A] from PHStkDaysStm" & WhYmd
RunqC "Select Stm,TarStkMths    into [#B] from PHTarMthsStm" & WhCo
RunqC "Update [#StkHldStm] x inner join [#A] a on a.Stm=x.Stm set x.StkDays=a.StkDays,x.RemSC=a.RemSC,x.StkMths=a.StkDays/12"
RunqC "Update [#StkHldStm] x inner join [#B] a on x.Stm=a.Stm set x.TarStkMths=a.TarStkMths"

'-- 6 Stp-CpyAddPHAtr
'   Oup-@StkHld{7}
MH.Ph7Tb.CpyAddAtr "#StkHld?", "@StkHld?"

'-- 7 Stp-AddF1F2
Dim I: For Each I In MH.PH.Itmy7
    RunqC "Alter Table [@StkHld" & I & "] add column F1 Text(1),F2 Text(1)"
Next

'-- 8 Stp-Reseq ===================================
MH.Ph7Tb.Rseq "@StkHld?", FF14Var

'-- 9 Stp-DrpTmp
If False Then
    DrpTTC "#A #B #YpSku #Piv #2Tot"
    DrpTnyC WPhTbny
End If
End Sub
Private Function WPhTbny() As String()
WPhTbny = MH.PH.Tmpy7Qtp("#StkHld?")
End Function
Private Sub WOStkHld_Add3_PhRollUpCol(SkuTbl$) 'Add 3-PHRollupCol with value to Tbl-@SkuT from qry-qSku_Main:qry which based on Tb-SKU and more Tbl.  more..
':3PHRollupCol: are 3 columns used to roll up to 6 lvl above of PH: They are {Stm BusArea PHL1..4}
':qSku_Main: is a query based on Tb-Sku with Sku as Pk and Tb-other and with at 3 fields-[Stm BusArea PHL4]
':SkuT: is a tbn with a fld-Sku.
RunqC FmtQQ("Alter Table [?] add column Stm Text(1),BusArea Text(4),PHL4 Text(10)", SkuTbl)
RunqC FmtQQ("Update [?] x inner join qSku_Main a on a.Sku=x.Sku set x.Stm=a.Stm,x.BusArea=a.BusArea,x.PHL4=a.PHL4", SkuTbl)
End Sub

Private Sub WOStkDays(Co As Byte)
'Oup:
'     @StkDaysStm: Stm         Stm
'     @StkDaysBus: Stm BusArea Stm PHSBus BusArea PHBus
'     @StkDaysL1 : Stm PHL1    Stm Srt1 PHNam
'     @StkDaysL2 : Stm PHL2    Stm Srt2 PHNam PHBrd
'     @StkDaysL3 : Stm PHL3    Stm Srt3 PHNam PHBrd PHQGp
'     @StkDaysL4 : Stm PHL4    Stm Srt4 PHNam PHBrd PHQGp PHQly
'     @StkDaysSku: Sku         Stm Srt4 PHNam PHBrd PHQGp PHQly Sku SkuDes
'Inp:
'     PHStkDaysStm = YY MM DD Co Stm         StkDays
'     PHStkDaysBus = YY MM DD Co Stm BusArea StkDays
'     PHStkDaysL1  = YY MM DD Co Stm PHL1    StkDays
'     PHStkDaysL2  = YY MM DD Co Stm PHL2    StkDays
'     PHStkDaysL3  = YY MM DD Co Stm PHL3    StkDays
'     PHStkDaysL4  = YY MM DD Co Stm PHL4    StkDays
'     PHStkDaysSku = YY MM DD Co Stm Sku     StkDays
'# Stp/Oup Des
'= ======= ===
'1 #Key    YY MM DD Co
'2 #D{7}
'3 #E{7}
'4 #OHSrc{7} #D: #Key YY MM DD Co
'5 Tmp#E
'     #EStm Stm     M01..15
'     #ESku Stm Sku M01..15
'     ..
'== Stp1-Key = N YY MM DD, Where *YY *MM and roll 15 back from given @Ymd and *DD is the max with the *YY-&-*MM ==============
'Oup:#Key         = Co N YY MM DD   ' DD is max(DD)
'Fm :PHStkDaysStm = YY MM DD Co Stm | ..
DrpPfxxC "#Key #D #E #OHSC"

RunqQQC "SELECT Top 15 Co,(YY-{0})*12+MM-{1} AS N, YY, MM, Max(x.DD) AS DDLng" & _
" INTO [#Key]" & _
" FROM PHStkDaysStm x" & _
" Where Co=" & Co & _
" Group BY Co,(YY-{0})*12+MM-{1},YY,MM" & _
" HAVING (YY-{0})*12+MM-{1} Between -14 And 0;", Myy, Mmm
    '== Adj DDLng to DD
    RunqC "Alter Table [#Key] add Column DD Byte"
    RunqC "Update [#Key] set DD = DDLng"
    RunqC "Alter Table [#Key] drop column DDLng"

'==Stp2-#D{7} ===============================================================================================
Const Sel$ = "Select x.N, x.YY, x.MM, x.DD, StkDays, RemSC, "
Const JnOn$ = " ON x.DD=a.DD AND x.MM=a.MM AND x.YY=a.YY and x.Co=a.Co;"
DoCmd.SetWarnings False
RunqC Sel & "a.Sku        INTO [#DSku] FROM [#Key] x INNER JOIN PHStkDaysSku a" & JnOn
RunqC Sel & "Stm          INTO [#DStm] FROM [#Key] x INNER JOIN PHStkDaysStm a" & JnOn
RunqC Sel & "Stm, BusArea INTO [#DBus] FROM [#Key] x INNER JOIN PHStkDaysBus a" & JnOn
RunqC Sel & "Stm, PHL1    INTO [#DL1]  FROM [#Key] x INNER JOIN PHStkDaysL1  a" & JnOn
RunqC Sel & "Stm, PHL2    INTO [#DL2]  FROM [#Key] x INNER JOIN PHStkDaysL2  a" & JnOn
RunqC Sel & "Stm, PHL3    INTO [#DL3]  FROM [#Key] x INNER JOIN PHStkDaysL3  a" & JnOn
RunqC Sel & "Stm, PHL4    INTO [#DL4]  FROM [#Key] x INNER JOIN PHStkDaysL4  a" & JnOn

'==Stp3-#E{7}================================================================================================
Const F15$ = "StkDays01 Integer, RemSC01 Double," & _
"StkDays02 Integer, RemSC02 Double," & _
"StkDays03 Integer, RemSC03 Double," & _
"StkDays04 Integer, RemSC04 Double," & _
"StkDays05 Integer, RemSC05 Double," & _
"StkDays06 Integer, RemSC06 Double," & _
"StkDays07 Integer, RemSC07 Double," & _
"StkDays08 Integer, RemSC08 Double," & _
"StkDays09 Integer, RemSC09 Double," & _
"StkDays10 Integer, RemSC10 Double," & _
"StkDays11 Integer, RemSC11 Double," & _
"StkDays12 Integer, RemSP12 Double," & _
"StkDays13 Integer, RemSC13 Double," & _
"StkDays14 Integer, RemSC14 Double," & _
"StkDays15 Integer, RemSC15 Double)"

RunqC "Create Table [#ESku] (Sku Text(20)," & F15
RunqC "INSERT INTO [#ESku] SELECT Distinct Sku From [#DSku]"
Dim N%: For N = 0 To -14 Step -1
Dim NStr$: NStr = Format(1 - N, "00")
RunqC "UPDATE [#ESku] x INNER JOIN [#DSku] a ON x.Sku=a.Sku SET x.StkDays" & NStr & "=a.StkDays,x.RemSC01=a.RemSC WHERE a.N=" & N
Next

'#EStm ..............................................................................................
DrpC "#EStm"
RunqC "Create Table [#EStm] (Stm Text(1)," & F15
RunqC "INSERT INTO [#EStm] SELECT Distinct Stm From [#DStm]"
For N = 0 To -14 Step -1
NStr = Format(1 - N, "00")
RunqC "UPDATE [#EStm] x INNER JOIN [#DStm] a ON x.Stm=a.Stm SET x.StkDays" & NStr & "=a.StkDays,x.RemSC01=a.RemSC WHERE a.N=" & N
Next

'#EBus ..............................................................................................
DrpC "#EBus"
RunqC "Create Table [#EBus] (Stm Text(1),BusArea Text(4)," & F15
RunqC "INSERT INTO [#EBus] SELECT Distinct Stm,BusArea From [#DBus]"
For N = 0 To -14 Step -1
NStr = Format(1 - N, "00")
RunqC "UPDATE [#EBus] x INNER JOIN [#Dbus] a ON x.Stm=a.Stm and x.BusArea=a.BusArea SET x.StkDays" & NStr & "=a.StkDays,x.RemSC01=a.RemSC WHERE a.N=" & N
Next

'#EL1 ..............................................................................................
DrpC "#EL1"
RunqC "Create Table [#EL1] (Stm Text(1),PHL1 Text(2)," & F15
RunqC "INSERT INTO [#EL1] SELECT Distinct Stm,PHL1 From [#DL1]"
For N = 0 To -14 Step -1
NStr = Format(1 - N, "00")
RunqC "UPDATE [#EL1] x INNER JOIN [#DL1] a ON x.Stm=a.Stm and x.PHL1=a.PHL1 SET x.StkDays" & NStr & "=a.StkDays,x.RemSC01=a.RemSC WHERE a.N=" & N
Next

'#EL2 ..............................................................................................
DrpC "#EL2"
RunqC "Create Table [#EL2] (Stm Text(1),PHL2 Text(4)," & F15
RunqC "INSERT INTO [#EL2] SELECT Distinct Stm,PHL2 From [#DL2]"
For N = 0 To -14 Step -1
NStr = Format(1 - N, "00")
RunqC "UPDATE [#EL2] x INNER JOIN [#DL2] a ON x.Stm=a.Stm and x.PHL2=a.PHL2 SET x.StkDays" & NStr & "=a.StkDays,x.RemSC01=a.RemSC WHERE a.N=" & N
Next

'#EL3 ..............................................................................................
DrpC "#EL3"
RunqC "Create Table [#EL3] (Stm Text(1),PHL3 Text(7)," & F15
RunqC "INSERT INTO [#EL3] SELECT Distinct Stm,PHL3 From [#DL3]"
For N = 0 To -14 Step -1
NStr = Format(1 - N, "00")
RunqC "UPDATE [#EL3] x INNER JOIN [#DL3] a ON x.Stm=a.Stm and x.PHL3=a.PHL3 SET x.StkDays" & NStr & "=a.StkDays,x.RemSC01=a.RemSC WHERE a.N=" & N
Next

'#EL4 ..............................................................................................
DrpC "#EL4"
RunqC "Create Table [#EL4] (Stm Text(1),PHL4 Text(10)," & F15
RunqC "INSERT INTO [#EL4] SELECT Distinct Stm,PHL4 From [#DL4]"
For N = 0 To -14 Step -1
NStr = Format(1 - N, "00")
RunqC "UPDATE [#EL4] x INNER JOIN [#DL4] a ON x.Stm=a.Stm and x.PHL4=a.PHL4 SET x.StkDays" & NStr & "=a.StkDays,x.RemSC01=a.RemSC WHERE a.N=" & N
Next


    '== Stp-OHSC{7} from @StkHld{7}
    '-- Create #OHSC{7} from @StkHld{7}
    RunqC "Select Sku           ,ScTot As SC into [#OHSCSku] from [@StkHldSku]"
    RunqC "Select Stm        ,ScTot As SC into [#OHSCStm] from [@StkHldStm] x inner join PHLStm a on a.Stream=x.Stream"
    RunqC "Select Stm,BusArea,ScTot As SC into [#OHSCBus] from [@StkHldBus] x inner join PHLStm a on a.Stream=x.Stream"
    RunqC "Select Stm,PHL1   ,ScTot As SC into [#OHSCL1]  from [@StkHldL1] x inner join PHLStm a on a.Stream=x.Stream"
    RunqC "Select Stm,PHL2   ,ScTot As SC into [#OHSCL2]  from [@StkHldL2] x inner join PHLStm a on a.Stream=x.Stream"
    RunqC "Select Stm,PHL3   ,ScTot As SC into [#OHSCL3]  from [@StkHldL3] x inner join PHLStm a on a.Stream=x.Stream"
    RunqC "Select Stm,PHL4   ,ScTot As SC into [#OHSCL4]  from [@StkHldL4] x inner join PHLStm a on a.Stream=x.Stream"
    
    'AddSC to #E{7} .............................................
    RunqC "Alter Table [#EStm] add column SC double"
    RunqC "Alter Table [#ESku] add column SC double"
    RunqC "Alter Table [#EBus] add column SC double"
    RunqC "Alter Table [#EL1] add column SC double"
    RunqC "Alter Table [#EL2] add column SC double"
    RunqC "Alter Table [#EL3] add column SC double"
    RunqC "Alter Table [#EL4] add column SC double"
    RunqC "Update [#ESku] x inner join [#OHSCSku] a on a.Sku=x.Sku                         set x.SC=a.SC"
    RunqC "Update [#EStm] x inner join [#OHSCStm] a on a.Stm=x.Stm                         set x.SC=a.SC"
    RunqC "Update [#EBus] x inner join [#OHSCBus] a on a.Stm=x.Stm and a.BusArea=x.BusArea set x.SC=a.SC"
    RunqC "Update [#EL1]  x inner join [#OHSCL1]  a on a.Stm=x.Stm and a.PHL1   =x.PHL1    set x.SC=a.SC"
    RunqC "Update [#EL2]  x inner join [#OHSCL2]  a on a.Stm=x.Stm and a.PHL2   =x.PHL2    set x.SC=a.SC"
    RunqC "Update [#EL3]  x inner join [#OHSCL3]  a on a.Stm=x.Stm and a.PHL3   =x.PHL3    set x.SC=a.SC"
    RunqC "Update [#EL4]  x inner join [#OHSCL4]  a on a.Stm=x.Stm and a.PHL4   =x.PHL4    set x.SC=a.SC"
    
    '== Stp-InsWithOHNoStkDays
    '   #E{7}->SC insert records where has OH, but no StkDays
    RunqC "Insert into [#ESku] (Sku        ,SC) select x.Sku          ,x.SC from [#OHScSku] x left join [#ESku] a on x.Sku=a.sku                   where a.Sku is null"
    RunqC "Insert into [#EL4]  (Stm,PHL4   ,SC) select x.Stm,x.PHL4   ,x.SC from [#OHScL4]  x left join [#EL4]  a on x.PHL4=a.PHL4 and x.Stm=a.Stm where a.Stm is null"
    RunqC "Insert into [#EL3]  (Stm,PHL3   ,SC) select x.Stm,x.PHL3   ,x.SC from [#OHScL3]  x left join [#EL3]  a on x.PHL3=a.PHL3 and x.Stm=a.Stm where a.Stm is null"
    RunqC "Insert into [#EL2]  (Stm,PHL2   ,SC) select x.Stm,x.PHL2   ,x.SC from [#OHScL2]  x left join [#EL2]  a on x.PHL2=a.PHL2 and x.Stm=a.Stm where a.Stm is null"
    RunqC "Insert into [#EL1]  (Stm,PHL1   ,SC) select x.Stm,x.PHL1   ,x.SC from [#OHScL1]  x left join [#EL1]  a on x.PHL1=a.PHL1 and x.Stm=a.Stm where a.Stm is null"
    RunqC "Insert into [#EBus] (Stm,BusArea,SC) select x.Stm,x.BusArea,x.SC from [#OHScBus] x left join [#EBus] a on x.BusArea=a.BusArea and x.Stm=a.Stm where a.Stm is null"

    '== Stp-AddAtr =============================================================================================
    '-- Add Attribute Fields
    MH.Ph7Tb.CpyAddAtr "#E?", "@StkDays?"

'== Stp-Reseq ======================================================================================================
Const RstFlds$ = "SC" & _
" StkDays01 RemSC01" & _
" StkDays02 RemSC02" & _
" StkDays03 RemSC03" & _
" StkDays04 RemSC04" & _
" StkDays05 RemSC05" & _
" StkDays06 RemSC06" & _
" StkDays07 RemSC07" & _
" StkDays08 RemSC08" & _
" StkDays09 RemSC09" & _
" StkDays10 RemSC10" & _
" StkDays11 RemSC11" & _
" StkDays12 RemSP12" & _
" StkDays13 RemSC13" & _
" StkDays14 RemSC14" & _
" StkDays15 RemSC15"
MH.Ph7Tb.Rseq "@StkDays?" ', RstFlds

'== Stp-DrpTmp ======================================================================================================
DrpPfxxC "#Key #D #E #OHSC"

'Dim I: For Each I In Split("Sku Stm Bus L1 L2 L3 L4")
'    RunqC "Drop Table [#D" & I & "]"
'    RunqC "Drop Table [#E" & I & "]"
'    RunqC "Drop Table [#OHSc" & I & "]"
'Next
'RunqC "Drop Table [#Key]"
End Sub
Private Function WCoYmd(Co As Byte) As CoYmd: WCoYmd = MHNw.CoYmd(Co, MYmd):    End Function
Private Function WCoYm(Co As Byte) As CoYm:    WCoYm = MHNw.CoYm(Co, Myy, Mmm): End Function
Private Sub B_WOFc():                                  WIniSamp: WOFc 86:       End Sub
Private Sub WOFc(Co As Byte)
'Aim: create @Fc{7} From FcSku
'Oup: @Fc{7}   {7LvlKey} | SC | M01..15

'## Stp       Oup      What
' 1 !@TmpFc     $Fc{7}   By calling FcTmp7_CoYM
' 2 !@TmpScOH   $ScOH{7} By Call ScOHTmp7CoYmd
' 3 !@AddScCol  $Fc{7}   add a column SC
' 4 !@InsOHNoSc $Fc{7}   add records from $ScOH for those with OH, but not Fc
'   !@AddSdRemCol $Fc{7} add two columns StkDays and RemSC to each lvl from PHStkDays{7}
' 5 !@AddAtr    @Fc{7}
' 6 !@DrpCoCol
'   !@ReSeq     @Fc{7}
' 7 !@DrpTmp    $Fc{7} $ScOH{7}
'== 1 !@TmpFc
Dim MCoYmd As CoYmd: MCoYmd = WCoYmd(Co)
MH.FcTmp7.CrtCoYm WCoYm(Co)

'== 2 !@TmpScOH
MH.SHldTmp7.ScOHCoYmd MCoYmd

'== 3 !@AddScCol
RunqC "Alter Table [$FcSku] Add Column SC Double"
RunqC "Alter Table [$FcL4] Add Column SC Double"
RunqC "Alter Table [$FcL3] Add Column SC Double"
RunqC "Alter Table [$FcL2] Add Column SC Double"
RunqC "Alter Table [$FcL1] Add Column SC Double"
RunqC "Alter Table [$FcBus] Add Column SC Double"
RunqC "Alter Table [$FcStm] Add Column SC Double"

'-- Update $Fc{7}->SC
RunqC "Update [$FcSku] x inner join [$ScOHSku] a on x.Co=a.Co and x.Sku=a.Sku         set x.SC = a.SC"
RunqC "Update [$FcL4]  x inner join [$ScOHL4]  a on x.Co=a.Co and x.PHL4=a.PHL4       set x.SC = a.SC"
RunqC "Update [$FcL3]  x inner join [$ScOHL3]  a on x.Co=a.Co and x.PHL3=a.PHL3       set x.SC = a.SC"
RunqC "Update [$FcL2]  x inner join [$ScOHL2]  a on x.Co=a.Co and x.PHL2=a.PHL2       set x.SC = a.SC"
RunqC "Update [$FcL1]  x inner join [$ScOHL1]  a on x.Co=a.Co and x.PHL1=a.PHL1       set x.SC = a.SC"
RunqC "Update [$FcBus] x inner join [$ScOHBus] a on x.Co=a.Co and x.BusArea=a.BusArea set x.SC = a.SC"
RunqC "Update [$FcStm] x inner join [$ScOHStm] a on x.Co=a.Co and x.Stm=a.Stm         set x.SC = a.SC"
'== 4 !@InsOHNoFc
RunqC "Insert into [$FcLSku] (Co,Sku        ,SC) select x.Co,x.Sku          ,x.SC from [$ScOHSku] x left join [$FcLSku] a on x.Co=a.Co and x.Sku=a.sku                   where a.Co is null"
RunqC "Insert into [$FcL4]  (Co,Stm,PHL4   ,SC) select x.Co,x.Stm,x.PHL4   ,x.SC from [$ScOHL4]  x left join [$FcL4]  a on x.Co=a.Co and x.PHL4=a.PHL4 and x.Stm=a.Stm where a.Co is null"
RunqC "Insert into [$FcL3]  (Co,Stm,PHL3   ,SC) select x.Co,x.Stm,x.PHL3   ,x.SC from [$ScOHL3]  x left join [$FcL3]  a on x.Co=a.Co and x.PHL3=a.PHL3 and x.Stm=a.Stm where a.Co is null"
RunqC "Insert into [$FcL2]  (Co,Stm,PHL2   ,SC) select x.Co,x.Stm,x.PHL2   ,x.SC from [$ScOHL2]  x left join [$FcL2]  a on x.Co=a.Co and x.PHL2=a.PHL2 and x.Stm=a.Stm where a.Co is null"
RunqC "Insert into [$FcL1]  (Co,Stm,PHL1   ,SC) select x.Co,x.Stm,x.PHL1   ,x.SC from [$ScOHL1]  x left join [$FcL1]  a on x.Co=a.Co and x.PHL1=a.PHL1 and x.Stm=a.Stm where a.Co is null"
RunqC "Insert into [$FcLBus] (Co,Stm,BusArea,SC) select x.Co,x.Stm,x.BusArea,x.SC from [$ScOHBus] x left join [$FcBus] a on x.Co=a.Co and x.BusArea=a.BusArea and x.Stm=a.Stm where a.Co is null"

WOStkHld_AddCol_StkDays_and_RemSc MCoYmd

'== 5 !@AddAtr
MH.Ph7Tb.CpyAddAtr "$Fc?", "@Fc?"

'== 6 !@DrpCoCol
Dim I: For Each I In WPhItmy7
    RunqC "alter Table [@Fc" & I & "] drop column Co"
Next
RunqC "Alter Table [@FcLSku] drop column Stm"
'== 7 !@Reseq
MH.Ph7Tb.Rseq "@Fc?" ', "SC StkDays RemSC" & _
" M01 M02 M03" & _
" M04 M05 M06" & _
" M07 M08 M09" & _
" M10 M11 M12" & _
" M13 M14 M15"

'== !@DrpTmp
MH.PH.DrpTmpy7Days
MH.PH.DrpTmpy7ScOH
End Sub
Private Function WPhItmy7() As String(): WPhItmy7 = MH.PH.Itmy7: End Function
Private Sub WOStkHld_AddCol_StkDays_and_RemSc(A As CoYmd)
'Dmp NyQtp2("RunqC ""Alter Table [$Fc?] add Column StkDays Integer, RemSC double""", Ph7Ss)
RunqC "Alter Table [$FcStm] add Column StkDays Integer, RemSC double"
RunqC "Alter Table [$FcBus] add Column StkDays Integer, RemSC double"
RunqC "Alter Table [$FcL1] add Column StkDays Integer, RemSC double"
RunqC "Alter Table [$FcL2] add Column StkDays Integer, RemSC double"
RunqC "Alter Table [$FcL3] add Column StkDays Integer, RemSC double"
RunqC "Alter Table [$FcL4] add Column StkDays Integer, RemSC double"
RunqC "Alter Table [$FcLSku] add Column StkDays Integer, RemSC double"

'Dmp NyQtp2("RunqC ""Alter Table [$Fc?] add Column StkDays Integer, RemSC double""", Ph7Ss)
Dim Jn$(6)
    Jn(0) = "x.Stm=a.Stm"
    Jn(1) = "x.Stm=a.Stm and x.BusArea=a.BusArea"
    Jn(2) = "x.Stm=a.Stm and x.PHL1=a.PHL1"
    Jn(3) = "x.Stm=a.Stm and x.PHL2=a.PHL2"
    Jn(4) = "x.Stm=a.Stm and x.PHL3=a.PHL3"
    Jn(5) = "x.Stm=a.Stm and x.PHL4=a.PHL4"
    Jn(6) = "x.Sku=a.Sku"
Dim K$(6)
    K(0) = "x.Stm"
    K(1) = "x.Stm,x.BusArea"
    K(2) = "x.Stm,x.PHL1"
    K(3) = "x.Stm,x.PHL2"
    K(4) = "x.Stm,x.PHL3"
    K(5) = "x.Stm,x.PHL4"
    K(6) = "x.Sku"

Dim Wh$: Wh = MH.TbOH.WhCoYmd(A)
Dim I, J%: For Each I In WPhItmy7
    DrpC "#A"
    RunqC FmtQQ("Select ?,StkDays,RemSC Into [#A] from [PHStkDaysFc?] ?", K(J), I, Wh)
    RunqC FmtQQ("Update [$Fc?] x inner Join [#A] a on ? Set x.StkDays=a.StkDays,x.RemSC=a.RemSC", I, Jn(J), K(J))
    J = J + 1
Next
DrpC "#A"
End Sub
Private Sub W_GenFxo(Co As Byte)
Dim Fxo$: Fxo = WFxoCo(Co)
Dim Tp$: Tp = MH.SHldTp.Tp
Dim B As Workbook: Set B = MXls.Workbooks(Fxo)
CpyFfn Tp, Fxo
Set B = MXls.Workbooks.Open(Fxo)
RfhLoWbRsC B
WFB__FmtWb B, Co
B.Save
B.Close
End Sub
Private Sub WMinv(): Minv MXls: End Sub
Private Function WFxoCo$(Co As Byte)
Select Case Co
Case 86: WFxoCo = MFxo86
Case 87: WFxoCo = MFxo87
End Select
End Function
Private Sub WFB__FmtWb(B As Workbook, Co As Byte)
WMinv ' This is need because, Merge & Unmerge will break under MiniState
MH.SHldFmtDteTit15Mth.Fmt B, MYm
WFB_FmtA1 B, Co
End Sub
Private Sub WFB_FmtA1(B As Workbook, Co As Byte)
Dim S As Worksheet: For Each S In B
    If MH.SHldWsn.IsSHld(S.Name) Or S.Name = "Index" Then
        WFB_SetAt S.Range("A1"), Co
    End If
Next
Set S = B.Sheets("Index")
S.Range("C2").Value = Now
End Sub
Private Function WFB_SfxA1$(Co As Byte)
Dim Con$: Con = Fun.NmCo(Co)
WFB_SfxA1 = " As At " & HypymdYmd(MYmd) & " (" & Con & ")"
End Function
Private Sub WFB_SetAt(A1 As Range, Co As Byte)
Dim A1Old$: A1Old = A1.Value
Dim A1Nw$: A1Nw = BefOrAll(A1Old, " As At") & WFB_SfxA1(Co)
A1.Value = A1Nw
End Sub
