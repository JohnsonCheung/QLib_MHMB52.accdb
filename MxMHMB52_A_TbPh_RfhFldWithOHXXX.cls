VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_TbPh_RfhFldWithOHXXX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Sub RfhFldWithOHXxx()
DoCmd.SetWarnings False
RunqC "Select Distinct Sku into [#OHSkuHst] from OH"
RunqC "Select Distinct Sku into [#OHSkuCur] from OH" & MH.TbOH.WhLas

RunqC "Select Distinct ProdHierarchy as PH into [#OHPHCur] from SKU where SKU in (Select Sku From [#OHSkuCur])"
RunqC "Select Distinct ProdHierarchy as PH into [#OHPHHst] from SKU where SKU in (Select Sku From [#OHSkuHst])"
RunqC "Select Distinct CByte(4) as Lvl, Left(PH,10) as PHL4 into [#PHL4Cur] from [#OHPHCur]"
RunqC "Select Distinct CByte(4) as Lvl, Left(PH,10) as PHL4 into [#PHL4Hst] from [#OHPHHst]"
RunqC "Select Distinct CByte(3) as Lvl, Left(PH,7) as PHL3 into [#PHL3Cur] from [#OHPHCur]"
RunqC "Select Distinct CByte(3) as Lvl, Left(PH,7) as PHL3 into [#PHL3Hst] from [#OHPHHst]"
RunqC "Select Distinct CByte(2) as Lvl, Left(PHL3,4) as PHL2 into [#PHL2Cur] from [#PHL3Cur]"
RunqC "Select Distinct CByte(2) as Lvl, Left(PHL3,4) as PHL2 into [#PHL2Hst] from [#PHL3Hst]"
RunqC "Select Distinct CByte(1) as Lvl, Left(PHL2,2) as PHL1 into [#PHL1Cur] from [#PHL2Cur]"
RunqC "Select Distinct CByte(1) as Lvl, Left(PHL2,2) as PHL1 into [#PHL1Hst] from [#PHL2Hst]"

RunqC "Update ProdHierarchy x Left Join [#OHPHCur] a on x.PH=a.PH set WithOHCur=Not IsNull(a.PH) where Lvl=4"
RunqC "Update ProdHierarchy x Left Join [#OHPHHst] a on x.PH=a.PH set WithOHHst=Not IsNull(a.PH) where Lvl=4"

RunqC "Update ProdHierarchy x inner Join [#PHL4Cur] a on x.Lvl=a.Lvl and x.PH=a.PHL4 set WithOHCur=Not IsNull(a.PHL4)"
RunqC "Update ProdHierarchy x inner Join [#PHL4Hst] a on x.Lvl=a.Lvl and x.PH=a.PHL4 set WithOHHst=Not IsNull(a.PHL4)"

RunqC "Update ProdHierarchy x inner Join [#PHL3Cur] a on x.Lvl=a.Lvl and x.PH=a.PHL3 set WithOHCur=Not IsNull(a.PHL3)"
RunqC "Update ProdHierarchy x inner Join [#PHL3Hst] a on x.Lvl=a.Lvl and x.PH=a.PHL3 set WithOHHst=Not IsNull(a.PHL3)"

RunqC "Update ProdHierarchy x inner Join [#PHL2Cur] a on x.Lvl=a.Lvl and x.PH=a.PHL2 set WithOHCur=Not IsNull(a.PHL2)"
RunqC "Update ProdHierarchy x inner Join [#PHL2Hst] a on x.Lvl=a.Lvl and x.PH=a.PHL2 set WithOHHst=Not IsNull(a.PHL2)"

RunqC "Update ProdHierarchy x inner Join [#PHL1Cur] a on x.Lvl=a.Lvl and x.PH=a.PHL1 set WithOHCur=Not IsNull(a.PHL1)"
RunqC "Update ProdHierarchy x inner Join [#PHL1Hst] a on x.Lvl=a.Lvl and x.PH=a.PHL1 set WithOHHst=Not IsNull(a.PHL1)"

RunqC "Drop Table [#OHSkuHst]"
RunqC "Drop Table [#OHSkuCur]"
RunqC "Drop Table [#OHPHHst]"
RunqC "Drop Table [#OHPHCur]"
RunqC "Drop Table [#PHL1Cur]"
RunqC "Drop Table [#PHL1Hst]"
RunqC "Drop Table [#PHL2Cur]"
RunqC "Drop Table [#PHL2Hst]"
RunqC "Drop Table [#PHL3Cur]"
RunqC "Drop Table [#PHL3Hst]"
RunqC "Drop Table [#PHL4Cur]"
RunqC "Drop Table [#PHL4Hst]"

End Sub
