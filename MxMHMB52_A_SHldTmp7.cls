VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_SHldTmp7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Sub ScOHDrp(): MH.PH.DrpTmpy7ScOH: End Sub
Private Sub B_TmpScOH()
ScOHYmd MH.TbOH.YmdLas
BrwTnyC MH.PH.Tmpy7Qtp("$ScOH?")
End Sub
Sub ScOHYmd(A As Ymd):     WScOH MH.TbOH.WhYmd(A):   End Sub 'Called by UpdTbPHStkDays7_FldStkDays_andRemSc
Sub ScOHCoYmd(A As CoYmd): WScOH MH.TbOH.WhCoYmd(A): End Sub

Private Sub WScOH(Wh$) 'Crt #ScOH{7} by @Wh StdCas OH
DoCmd.SetWarnings False
'Oup: $ScOH{7}
'Inp: OH = YY MM DD Co Sku | Btl
'Ref: Sku     = Sku [Btl/SC]
'Ref:
'Oup: $ScOHSku  Co Stm Sku  | SC
'Oup: $ScOHL4   Co Stm PHL4 | SC
'Oup: $ScOHL3   Co Stm PHL3 | SC
'Oup: $ScOHL2   Co Stm PHL2 | SC
'Oup: $ScOHL1   Co Stm PHL1 | SC
'Oup: $ScOHBus  Co Stm BusArea | SC
'Oup: $ScOHStm  Co Stm      | SC

'$ScOHSku
RunqC FmtQQ("Select Distinct Co,Sku,Sum(x.Btl) as Btl,CDbl(0) As SC into [$ScOHSku] from OH x ? Group by Co,Sku", Wh)
RunqC "Alter Table [$ScOHSku] add Column Stm Text(1)"
RunqC "Update [$ScOHSku] x inner join qSku_Main a on x.Sku=a.Sku set SC = Btl/[Btl/SC],x.Stm=a.Stm"
RunqC "Alter Table [$ScOHSku] Drop Column Btl"

'Oup: $ScOHL4
'Fm : $ScOHSku
'Tmp: #A
'Ref: qSku_Main      -> Sku PLH4
RunqC "Select Co,x.Sku,PHL4,SC Into [#A] from [$ScOHSku] a left join qSku_Main x on a.SKu=x.SKu"
RunqC "Select Distinct x.Co,Stm,PHL4,Sum(x.SC) as SC" & _
" into [$ScOHL4]" & _
" from [#A] x" & _
" left join [$ScOHSku] a on a.Sku=x.Sku" & _
" group by x.Co,Stm,PHL4"
RunqC "Drop Table [#A]"

'Oup: $ScOHL3  | Fm : $ScOHL4
'Oup: $ScOHL2  | Fm : $ScOHL3
'Oup: $ScOHL1  | Fm : $ScOHL2
'Oup: $ScOHStm | Fm : $ScOHL1
RunqC "Select Distinct Co,Stm,Left(PHL4,7) as PHL3,Sum(x.SC) as SC into [$ScOHL3] from [$ScOHL4] x Group By Co,Stm,Left(PHL4,7)"
RunqC "Select Distinct Co,Stm,Left(PHL3,4) as PHL2,Sum(x.SC) as SC into [$ScOHL2] from [$ScOHL3] x Group By Co,Stm,Left(PHL3,4)"
RunqC "Select Distinct Co,Stm,Left(PHL2,2) as PHL1,Sum(x.SC) as SC into [$ScOHL1] from [$ScOHL2] x Group By Co,Stm,Left(PHL2,2)"
RunqC "Select Distinct Co,Stm,Sum(x.SC) as SC into [$ScOHStm] from [$ScOHL1] x Group by Co,Stm"

'Oup: $ScOHBus | Fm : $ScOHSku
'Ref: Sku_Main  = Sku Stm BusArea
RunqC "Select Distinct Co,Stm,BusArea,Sum(x.SC) as SC" & _
" into [$ScOHBus]" & _
" from [$ScOHSku] x left join qSku_Main a on a.Sku=x.Sku" & _
" Group By Co,Stm,BusArea"
End Sub
