VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_SHldStkDaysCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Private Type StkDaysRem
    StkDays As Integer
    RemSC As Double  ' See Subr-StkDaysRem
End Type

Private Sub W_TmpDays7__T7__Tst():                 W_TmpDays7__T7 WWSampYmd:     End Sub
Private Sub W_UpdTblPHStkDays7__Tst():             W_UpdTblPHStkDays7 WWSampYmd: End Sub
Private Sub Calc__Tst():                           Calc WWSampYmd:               End Sub
Private Function WWSampYmd() As Ymd:   WWSampYmd = Ymd(19, 12, 24):              End Function
Sub Calc(A As Ymd) 'Upd Tb-PHStkDays{7}->StkDays/RemSc
'Inp: OH
'Inp: FcSku = VerYY VerMM Co Stm Sku M01..15
'Where {7} of PHStkDays{7} = {Sku Stm Bus L1..4}
'---
Sts "Update Stock Days ....":
MH.SHldTmp7.ScOHYmd A        '$ScOH{7} are created        = Co Stm {7} SC
MH.FcTmp7.CrtYm YmYmd(A)    '$Fc{7} are created        = Co Stm {7} M01..M15
W_TmpDays7__T7 A     '$Days{7} are created      = Co Stm {7} StkDays
W_UpdTblPHStkDays7 A     'Load $Days{7} into PHStkDays{7} = Y M D Co Stm {7} StkDays
MH.FcTmp7.Drp
MH.SHldTmp7.ScOHDrp
MH.PH.DrpTmpy7Days
End Sub

Private Sub W_TmpDays7__T7(A As Ymd)
'Inp: $ScOH{7} = Co Stm {7} SC
'Inp: $Fc{7} = Co Stm {7} M01..15
'Oup: $Days{7} = Co Stm {7} StkDays
'---
'#1 Tmp: #Fc{7}OH = Co Stm {7} M01..15 SC           ! SC is added at end of $Fc{7} to become #Fc{7}X
'#2 Tmp: $Fc{7}OH = ..                    StkDays   ! Add StkDays.  Each Record has enought data to calc StkDays
'#3 Oup: $Days{7} = Co Stm {7} StkDays              ! From $Fc{7}OH, just remove the M01..M15
'-----
'== Stp-TmpFcXOH: $Fc{7}OH
'      Fm : $Fc{7}
Const Sel$ = "Select x.*," & _
"a.M01,a.M02,a.M03," & _
"a.M04,a.M05,a.M06," & _
"a.M07,a.M08,a.M09," & _
"a.M10,a.M11,a.M12," & _
"a.M13,a.M14,a.M15,CInt(0) as StkDays,CDbl(0) as RemSC"
RunqC Sel & " into [#FcSkuOH] from [$ScOHSku] x left join [$FcSku] a on a.Co=x.Co and a.Stm=x.Stm and a.Sku=x.Sku"
RunqC Sel & " into [#FcBusOH] from [$ScOHBus] x left join [$FcBus] a on a.Co=x.Co and a.Stm=x.Stm and a.BusArea=x.BusArea"
RunqC Sel & " into [#FcStmOH] from [$ScOHStm] x left join [$FcStm] a on a.Co=x.Co and a.Stm=x.Stm"
RunqC Sel & " into [#FcL1OH]  from [$ScOHL1]  x left join [$FcL1]  a on a.Co=x.Co and a.Stm=x.Stm and a.PHL1=x.PHL1"
RunqC Sel & " into [#FcL2OH]  from [$ScOHL2]  x left join [$FcL2]  a on a.Co=x.Co and a.Stm=x.Stm and a.PHL2=x.PHL2"
RunqC Sel & " into [#FcL3OH]  from [$ScOHL3]  x left join [$FcL3]  a on a.Co=x.Co and a.Stm=x.Stm and a.PHL3=x.PHL3"
RunqC Sel & " into [#FcL4OH]  from [$ScOHL4]  x left join [$FcL4]  a on a.Co=x.Co and a.Stm=x.Stm and a.PHL4=x.PHL4"
'
'== Stp-UpdStkDays $Fc{7}OH: Update StkDays & RemSC
WT7_UpdTblT_FldStkDaysRem "#FcSkuOH", A
Stop
WT7_UpdTblT_FldStkDaysRem "#FcStmOH", A
WT7_UpdTblT_FldStkDaysRem "#FcL1OH", A
WT7_UpdTblT_FldStkDaysRem "#FcL2OH", A
WT7_UpdTblT_FldStkDaysRem "#FcL3OH", A
WT7_UpdTblT_FldStkDaysRem "#FcL4OH", A
WT7_UpdTblT_FldStkDaysRem "#FcBusOH", A
'== Stp-Days Create: $Days{7} from #Fc{7}OH just drop all the M01..15
WT7_CrtTmpDaysAndDrpM1To15 "Sku"  '$DaysSku is created
WT7_CrtTmpDaysAndDrpM1To15 "Stm"
WT7_CrtTmpDaysAndDrpM1To15 "Bus"
WT7_CrtTmpDaysAndDrpM1To15 "L1"
WT7_CrtTmpDaysAndDrpM1To15 "L2"
WT7_CrtTmpDaysAndDrpM1To15 "L3"
WT7_CrtTmpDaysAndDrpM1To15 "L4"

'---=
RunqC "Drop Table [#FcSkuOH]"
RunqC "Drop Table [#FcStmOH]"
RunqC "Drop Table [#FcBusOH]"
RunqC "Drop Table [#FcL1OH]"
RunqC "Drop Table [#FcL2OH]"
RunqC "Drop Table [#FcL3OH]"
RunqC "Drop Table [#FcL4OH]"
End Sub
Private Sub WT7_CrtTmpDaysAndDrpM1To15(LvlItm$)
RunqC "Select * into [$Days" & LvlItm & "] from [#Fc" & LvlItm & "OH]"
RunqC "Alter Table [$Days" & LvlItm & "] drop column " & _
"M01,M06,M11," & _
"M02,M07,M12," & _
"M03,M08,M13," & _
"M04,M09,M14," & _
"M05,M10,M15"
End Sub
Private Sub WT7_FcAy__Tst()
Dim Rs As Dao.Recordset: Set Rs = CurrentDb.TableDefs("FcSku").OpenRecordset
With Rs
    Dim F!: F = FactorRemDays(Now)
    Dim Fc#()
    While Not .EOF
        Fc = WT7_FcAy(Rs, F)
        Dim J%: For J = 0 To Si(Fc) - 1
            Debug.Print Fc(J);
        Next
        Debug.Print
        .MoveNext
    Wend
End With
End Sub
Private Function WT7_FcAy(Rs As Dao.Recordset, FstMthRemDaysFac!) As Double()
'@Rs: It has M01..15
'Return : the WT7_FcAy with FstMth adjust to remaining days as in @FstMthRemDaysFac
'         Trim all the end element if it is zero
Dim OFc#(): ReDim OFc(14)
Dim J%: For J = 0 To 14
    OFc(J) = Nz(Rs.Fields("M" & Format(J + 1, "00")).Value, 0)
Next
OFc(0) = OFc(0) * FstMthRemDaysFac ' Adjust the first Month
For J = 14 To 0 Step -1
    If OFc(J) <> 0 Then
        ReDim Preserve OFc(J)
        WT7_FcAy = OFc
        Exit Function
    End If
Next
'-- All Fc is Zero, just return
End Function

Private Sub WT7_UpdTblT_FldStkDaysRem(T$, A As Ymd)
'@T: :#Fc{}OH: StkDays RemSC SC M01..15
'Oup: @T->StkDays & RemSC are updated
Dim Y As Byte, M As Byte: Y = A.Y: M = A.M
Dim F!:    F = FactorRemDays(DteYmd(A))
Dim D() As Byte: D = WWDayyYmd(A)
Dim Rs As Dao.Recordset: Set Rs = RsTblC(T)
With Rs
    Dim Fc#(), SC#
    While Not .EOF
        Fc = WT7_FcAy(Rs, F)
        SC = !SC
        Dim C As StkDaysRem: C = WT7_StkDaysRem(SC, Fc, D)
        .Edit
            !StkDays = C.StkDays
            !RemSC = C.RemSC
        .Update
        .MoveNext
    Wend
End With
End Sub
Private Sub WWDayyYmd__Tst(): DmpAy WWDayyYmd(Ymd(19, 11, 28)): End Sub
Private Function WWDayyYmd(A As Ymd) As Byte()
'Return : 15 month's days with @@Days(0) is the remaining date of the month of @A
Dim O() As Byte: ReDim O(14)
Dim D As Date: D = DteYmd(A)
O(0) = RemDays(D)
Dim J%: For J = 1 To 14
    D = DteFstNxtMth(D)
    O(J) = NDay(D)
Next
WWDayyYmd = O
End Function

Private Sub W_UpdTblPHStkDays7(A As Ymd)
'Upd: PHStkDays{7}
'By : $Days{7}
DoCmd.SetWarnings False
Dim W$: W = MH.TbOH.WhYmd(A)
RunqC "Delete * from PHStkDaysSku" & W
RunqC "Delete * from PHStkDaysBus" & W
RunqC "Delete * from PHStkDaysStm" & W
RunqC "Delete * from PHStkDaysL1" & W
RunqC "Delete * from PHStkDaysL2" & W
RunqC "Delete * from PHStkDaysL3" & W
RunqC "Delete * from PHStkDaysL4" & W
Dim Y As Byte, M As Byte, D As Byte
With A
    Y = .Y
    M = .M
    D = .D
End With
RunqC FmtQQ("Insert into PHStkDaysSku (YY, MM, DD,Co,Stm,Sku,    StkDays,RemSC) Select ?,?,?,Co,Stm,Sku,    StkDays,RemSC from [$DaysSku]", Y, M, D)
RunqC FmtQQ("Insert into PHStkDaysStm (YY, MM, DD,Co,Stm,        StkDays,RemSC) Select ?,?,?,Co,Stm,        StkDays,RemSC from [$DaysStm]", Y, M, D)
RunqC FmtQQ("Insert into PHStkDaysBus (YY, MM, DD,Co,Stm,BusArea,StkDays,RemSC) Select ?,?,?,Co,Stm,BusArea,StkDays,RemSC from [$DaysBus]", Y, M, D)
RunqC FmtQQ("Insert into PHStkDaysL1  (YY, MM, DD,Co,Stm,PHL1,   StkDays,RemSC) Select ?,?,?,Co,Stm,PHL1,   StkDays,RemSC from [$DaysL1]", Y, M, D)
RunqC FmtQQ("Insert into PHStkDaysL2  (YY, MM, DD,Co,Stm,PHL2,   StkDays,RemSC) Select ?,?,?,Co,Stm,PHL2,   StkDays,RemSC from [$DaysL2]", Y, M, D)
RunqC FmtQQ("Insert into PHStkDaysL3  (YY, MM, DD,Co,Stm,PHL3,   StkDays,RemSC) Select ?,?,?,Co,Stm,PHL3,   StkDays,RemSC from [$DaysL3]", Y, M, D)
RunqC FmtQQ("Insert into PHStkDaysL4  (YY, MM, DD,Co,Stm,PHL4,   StkDays,RemSC) Select ?,?,?,Co,Stm,PHL4,   StkDays,RemSC from [$DaysL4]", Y, M, D)
End Sub
Private Function WT7_StkDaysRem(SC#, Fc#(), Days() As Byte) As StkDaysRem
'@Fc :SCAy #Forecast-of-Each-Month-in-StdCase#
'@SC :SC   #OH-in-SC#
'@Days     #Days-of-each-Month# ! @Days(0) has adjusted according the Given Ymd
':StkDays: :Days ! Number of days can cover the Fc-Quantity for each months.
':RemSC:   :Dbl  ! #Remaing-SC#
'                ! It may in 1 of 3 conditions:
'               !   1. If All @SC can be consumed:                    RemSc <= 0
'               !   2. No Forecast, that means zero element in @Fc:   RemSc <= -1
'               !   3. {Days}+     (THe @SC is not                    RemSc <= A positive number
'Return :StkDays !  {SC#} (StdCase) by {Fc#()} and {Days}
If Si(Fc) = 0 Then
    WT7_StkDaysRem.StkDays = 9999
    WT7_StkDaysRem.RemSC = SC
    Exit Function
End If
Dim RemSC#: RemSC = SC
Dim ODays#
Dim J%: For J = 0 To UBound(Fc)
    If RemSC < Fc(J) Then WT7_StkDaysRem.StkDays = Round(ODays + Days(J) * RemSC / Fc(J)): Exit Function
    RemSC = RemSC - Fc(J)
    ODays = ODays + Days(J)
Next
If RemSC > 0 Then
    WT7_StkDaysRem.StkDays = Round(ODays)
    WT7_StkDaysRem.RemSC = RemSC
    Exit Function
End If
WT7_StkDaysRem.StkDays = Round(ODays)
End Function
