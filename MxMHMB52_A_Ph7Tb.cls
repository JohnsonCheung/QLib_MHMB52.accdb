VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_Ph7Tb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_Ph_Fun_Ph7Tb."
Private Sub B_WSqy(): D WSqy("#Fm?", "#To?"): End Sub
Private Sub B_CpyAddAtr()
MH.FcTmp7.CrtLasYm
MH.PhTmp5.Drp
DrpPfxxC "#Fc"
CpyAddAtr "$Fc?", "#Fc?": BrwTbQtpC "#Fc"
MH.FcTmp7.Drp
MH.PhTmp5.Drp
End Sub

Sub RplStmByStream(QtpTbn$) ' Replace Ph7Tbl->Stm by Stream
Const CSub$ = CMod & "RplStmByStream"
Dim T: For Each T In NyQtp2(QtpTbn, MH.PH.Ss7)
    GoSub ChkFld
    GoSub AddStream
    GoSub UpdStream
    GoSub DrpStm
Next
Exit Sub
ChkFld:
    ChkFldExist CDb, T, "Stm", CSub
    ChkFldNExi CDb, T, "Stream", CSub
    Return
AddStream:
    RunqC FmtQQ("Alter Table [?] add column Stream Text(6)", T)
    Return
UpdStream:
    RunqC FmtQQ("Update [?] x inner join PHLStm a on a.Stm=x.Stm set x.Stream=a.Stream", T)
    Return
DrpStm:
    RunqC FmtQQ("Alter Table [?] drop column Stm", T)
    Return
End Sub
Sub Rseq(QtpTbn$, Optional FfRst$) ' RseqFld for 7-QtpTbn-Ph7Ss tables by each of their PHAtrFlds
Const FfPHStm = "Stream"
Const FfPHBus$ = "Stream SrtBus BusArea PHBus"
Const FfPHL1$ = "Stream  Srt1   PHL1    PHNam"
Const FfPHL2$ = "Stream  Srt2   PHL2    PHNam PHBrd"
Const FfPHL3$ = "Stream  Srt3   PHL3    PHNam PHBrd PHQGp"
Const FfPHL4$ = "Stream  Srt4   PHL4    PHNam PHBrd PHQGp PHQly"
Const FfPHSku$ = "Stream Srt4   Sku     SkuDes PHL4 PHNam PHBrd PHQGp PHQly BusArea PHBus SrtBus"
Dim FF$: FF = NBPfxSpc(FfRst)
RseqFldC MH.PhTmpn.LStm(QtpTbn), FfPHStm & FF
RseqFldC MH.PhTmpn.LBus(QtpTbn), FfPHBus & FF
RseqFldC MH.PhTmpn.L1(QtpTbn), FfPHL1 & FF
RseqFldC MH.PhTmpn.L2(QtpTbn), FfPHL2 & FF
RseqFldC MH.PhTmpn.L3(QtpTbn), FfPHL3 & FF
RseqFldC MH.PhTmpn.L4(QtpTbn), FfPHL4 & FF
RseqFldC MH.PhTmpn.LSku(QtpTbn), FfPHSku & FF
End Sub

Sub CpyAddAtr(QtpFm$, QtpTo$) 'Crt-Tables-by-copying(Cpy)-(7)-tables-(And)-(Add)ing-(PHAtr)-fields
'Ref: 5-$PHL*-Tables: $PHL1..4 & $PHLSku(See !TmpPH).  Assumed they are existed.  That means calling this Sub, !TmpPH must be created first
'Inp: 7FmTbl : as identified by @QtpFm
'       Example: if QtpFm = "$FcStkDays_?"
'       tb-$FcStkDays_Stm should have field Stm
'       tb-$FcStkDays_Bus should have field Stm,BusArea
'       tb-$FcStkDays_Sku should have field Stm,Sku
'       tb-$FcStkDays_L1  should have field Stm,PHL1
'       tb-$FcStkDays_L2  should have field Stm,PHL2
'       tb-$FcStkDays_L3  should have field Stm,PHL3
'       tb-$FcStkDays_L4  should have field Stm,PHL4
'Oup: 7ToTbl as identified by @QtpTo
'       Example: if QtpTo = "@FcStkDays_?"
'                         Adding Column
'                         =============
'       tb-@FcStkDays_Stm no fld added
'       tb-@FcStkDays_Bus Srt  PHBus
'       tb-@FcStkDays_Sku Srt4 PHNam PHBrd PHQGp PHQly SkuDes
'       tb-@FcStkDays_L4  Srt4 PHNam PHBrd PHQGp PHQly
'       tb-@FcStkDays_L3  Srt3 PHNam PHBrd PHQGp
'       tb-@FcStkDays_L2  Srt2 PHNam PHBrd
'       tb-@FcStkDays_L1  Srt1 PHNam

StsQry FmtQQ("Cpy7andAddPHAtr: From[?] to[?]...", QtpFm, QtpTo)
DrpQtp2C QtpTo, MH.PH.Ss7
RunqSqyC WSqy(QtpFm, QtpTo)
'BrwTbQtpC QtpFm, Ph7Ss: Stop
'BrwTbQtpC QtpTo, Ph7Ss: Stop
End Sub
Private Function WSqy(QtpFm$, QtpTo$) As String()
Const SqlSku = "SELECT x.*,a.SkuDes,Srt4,PHL4,PHNam,PHBrd,PHQGp,PHQly,BusArea,PHBus,SrtBus" & _
                                                      " Into [?] From [?] x LEFT JOIN [$PHLSku] a on x.Sku    =a.Sku     Order by x.Stm,Srt4,x.Sku"
Const SqlStm = "SELECT x.*                              Into [?] From [?] x                                              Order by Stm"
Const SqlBus = "SELECT x.*,PHSBus as SrtBus,PHBus       Into [?] From [?] x LEFT JOIN PHLBus    a on x.BusArea=a.BusArea Order by x.Stm,PHSBus,PHBus"
Const SqlL4$ = "SELECT x.*,Srt4,PHNam,PHBrd,PHQGp,PHQly Into [?] From [?] x LEFT JOIN [$PHL4]   a on x.PHL4   =a.PHL4    Order by x.Stm,Srt4,PHNam,PHBrd,PHQGp,PHQly"
Const SqlL3$ = "SELECT x.*,Srt3,PHNam,PHBrd,PHQGp       Into [?] From [?] x LEFT JOIN [$PHL3]   a on x.PHL3   =a.PHL3    Order by x.Stm,Srt3,PHNam,PHBrd,PHQGp"
Const SqlL2$ = "SELECT x.*,Srt2,PHNam,PHBrd             Into [?] From [?] x LEFT JOIN [$PHL2]   a on x.PHL2   =a.PHL2    Order by x.Stm,Srt2,PHNam,PHBrd"
Const SqlL1$ = "SELECT x.*,Srt1,PHNam                   Into [?] From [?] x LEFT JOIN [$PHL1]   a on x.PHL1   =a.PHL1    Order By x.Stm,Srt1,PHNam"
Dim O$()
PushI O, FmtQQ(SqlSku, MH.PhTmpn.LSku(QtpTo), MH.PhTmpn.LSku(QtpFm))
PushI O, FmtQQ(SqlStm, MH.PhTmpn.LStm(QtpTo), MH.PhTmpn.LStm(QtpFm))
PushI O, FmtQQ(SqlBus, MH.PhTmpn.LBus(QtpTo), MH.PhTmpn.LBus(QtpFm))
PushI O, FmtQQ(SqlL1, MH.PhTmpn.L1(QtpTo), MH.PhTmpn.L1(QtpFm))
PushI O, FmtQQ(SqlL2, MH.PhTmpn.L2(QtpTo), MH.PhTmpn.L2(QtpFm))
PushI O, FmtQQ(SqlL3, MH.PhTmpn.L3(QtpTo), MH.PhTmpn.L3(QtpFm))
PushI O, FmtQQ(SqlL4, MH.PhTmpn.L4(QtpTo), MH.PhTmpn.L4(QtpFm))
WSqy = O
End Function
