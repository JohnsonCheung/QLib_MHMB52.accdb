VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_FcTmp7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_FcTmp7."
Function Tbnn$(): Tbnn = NnQtp2("$Fc?", MH.PH.Ss7): End Function
Function Tny() As String(): SplitSpc (Tbnn): End Function
Sub Brw():              BrwTnyC Tny:                  End Sub
Sub CrtYm(A As Ym):     CrtWh MH.FcWh.WhYm(A.Y, A.M): End Sub
Sub CrtCoYm(A As CoYm): CrtWh MH.FcWh.WhCoYm(A):      End Sub
Sub CrtLasYm():         CrtYm MH.FcLas.Ym:            End Sub
Private Sub CrtWh(Wh$)
'Oup: Create [$Fc? Stm Bus L1 L2 L3 L4 Sku] from FcSku & with referring [Sku BusArea ProdHierarchy]
'  FcLSku -> $FcLSku -> $FcL4 -> $FcL3 -> $FcL2 -> $FcL1 -> $FcStm
'                  -> $FcBus
'  $FcLSku = Co Stm Sku  *M15
'  $FcLBus = Co Stm Bus  *M15
'  $FcL4  = Co Stm PHL4 *M15
'  $FcL3  = Co Stm PHL3 *M15
'  $FcL2  = Co Stm PHL2 *M15
'  $FcL1  = Co Stm PHL1 *M15
'  $FcLStm = Co Stm      *M15
'       L1..4 = Co Stm PHL1..4 {Dta}
'       Sku   = Co Stm Sku     {Dta}
'           where {Dta} is M01..15
'     Select from FcSku into $FcSku and up to above 6 level
'     It is called by subr-Fc_Calc & subr-Fc_Export.
'Inp: FcSku
'Ref: Sku
StsQry "Forecast"
DoCmd.SetWarnings False
DrpPfxxC "#Sku $FcL"

Const Sum$ = _
"Sum(x.M01) As M01,Sum(x.M02) as M02," & _
"Sum(x.M03) As M03,Sum(x.M04) as M04," & _
"Sum(x.M05) As M05,Sum(x.M06) as M06," & _
"Sum(x.M07) As M07,Sum(x.M08) as M08," & _
"Sum(x.M09) As M09,Sum(x.M10) as M10," & _
"Sum(x.M11) As M11,Sum(x.M12) as M12," & _
"Sum(x.M13) As M13,Sum(x.M14) as M14," & _
"Sum(x.M15) As M15"

'== $FcSku <- FcSku
RunqC "select Distinct Co,Stm,Sku,M01,M02,M03,M04,M05,M06,M07,M08,M09,M10,M11,M12,M13,M14,M15" & _
" Into [$FcLSku]" & _
" from FcSku" & Wh
'== #Sku <- Sku
RunqC "Select Sku,BusArea,Left(ProdHierarchy,10) as PHL4 into [#Sku] from Sku"

'== $Fc{6} <- FcSku & #Sku
RunqC "select Distinct Co,Stm,BusArea,             " & Sum & " Into [$FcLBus] from [$FcLSku] x left join [#Sku] a on x.Sku=a.Sku group by Co,Stm,BusArea"
RunqC "Select Distinct Co,Stm,PHL4,                " & Sum & " Into [$FcL4]   from [$FcLSku] x left join [#Sku] a on x.Sku=a.Sku group by Co,Stm,PHL4"
RunqC "select Distinct Co,Stm,Left(PHL4,7) as PHL3," & Sum & " Into [$FcL3]   from [$FcL4] x group by Co,Stm,Left(PHL4,7)"
RunqC "select Distinct Co,Stm,Left(PHL3,4) as PHL2," & Sum & " Into [$FcL2]   from [$FcL3] x group by Co,Stm,Left(PHL3,4)"
RunqC "Select Distinct Co,Stm,Left(PHL2,2) as PHL1," & Sum & " Into [$FcL1]   from [$FcL2] x group by Co,Stm,Left(PHL2,2)"
RunqC "select Distinct Co,Stm,                     " & Sum & " Into [$FcLStm] from [$FcL1] x group by Co,Stm"

DrpC "#Sku"
End Sub

Sub Drp(): DrpTnyC Tny: End Sub
