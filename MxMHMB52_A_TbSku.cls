VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_TbSku"
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
RunqC "Update Sku x Left Join [#OHSkuCur] a on x.Sku=a.Sku set WithOHCur=Not IsNull(a.Sku)"
RunqC "Update Sku x Left Join [#OHSkuHst] a on x.Sku=a.Sku set WithOHHst=Not IsNull(a.Sku)"
RunqC "Drop Table [#OHSkuHst]"
RunqC "Drop Table [#OHSkuCur]"
End Sub

Sub RfhFldOvr() ' #Rfh-Tb-Sku-Override-Fields#
'BusArea
'BusAreaSap
'BusAreaOvr
'Litre/BtlSap
'Litre/BtlOvr
'Litre/Btl
RunqC "Update Sku Set BusArea='' Where BusArea Is Null"
RunqC "Update Sku Set BusAreaOvr='' Where BusAreaOvr Is Null"
RunqC "Update Sku Set BusAreaSap='' Where BusAreaSap Is Null"
RunqC "Update Sku Set [Litre/Btl]=0 Where [Litre/Btl] is null"
RunqC "Update Sku Set [Litre/BtlSap]=0 Where [Litre/BtlSap] is null"
RunqC "Update Sku Set [Litre/BtlOvr]=0 Where [Litre/BtlOvr] is null"

RunqC "Update Sku set BusArea=Trim(IIf(BusAreaOvr='',BusAreaSap,BusAreaOvr)) where BusArea<>Trim(IIf(BusAreaOvr='',BusAreaSap,BusAreaOvr))"
RunqC "Update Sku set [Litre/Btl]=IIf([Litre/BtlOvr]=0,[Litre/BtlSap],[Litre/BtlOvr]) where [Litre/Btl]<>IIf([Litre/BtlOvr]=0,[Litre/BtlSap],[Litre/BtlOvr])"
End Sub
