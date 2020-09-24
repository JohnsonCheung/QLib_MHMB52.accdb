VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A__MH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A__MH."
Public MHNw As New MxMHMB52_MHNw
Public MHRep As New MxMHMB52_MHRep
Public MHFun As New MxMHMB52_MHFun
Public MHBtnChk As New MxMHMB52_MHBtnChk
Public MHBtnRen As New MxMHMB52_MHBtnRen
Public FcIO As New MxMHMB52_A_FcIo
Public FcTp As New MxMHMB52_A_FcTp
Public TbOH As New MxMHMB52_A_TbOH
Public FcLas As New MxMHMB52_A_FcLas
Public FcFxwMH As New MxMHMB52_A_FcFxwMH
Public Skuy As New MxMHMB52_A_Skuy
Public TbSku As New MxMHMB52_A_TbSku
Public SkuLoad As New MxMHMB52_A_SkuLoad
Public FcLoad As New MxMHMB52_A_FcLoad
Public FcFxwUD As New MxMHMB52_A_FcFxwUD
Public MB52Rpt As New MxMHMB52_A_MB52Rpt
Public MB52Load As New MxMHMB52_A_MB52Load
Public FcRpt As New MxMHMB52_A_FcRpt
Public MB52IO As New MxMHMB52_A_MB52Io
Public MB52Las As New MxMHMB52_A_MB52Las
Public MB52Tp As New MxMHMB52_A_MB52Tp
Public TbReport As New MxMHMB52_A_TbReport
Public SHldTp As New MxMHMB52_A_SHldTp
Public FcWh As New MxMHMB52_A_FcWh
Public FcSamp As New MxMHMB52_A_FcSamp
Public FcTmp7 As New MxMHMB52_A_FcTmp7
Public TbFcFmPth As New MxMHMB52_A_TbFcFmPth
Public TbFcFmTbFcSku As New MxMHMB52_A_TbFcFmTbFcSku
Public GitIO As New MxMHMB52_A_GitIO
Public GitLoad As New MxMHMB52_A_GitLoad
Public PH As New MxMHMB52_A_PH
Public PHLoad As New MxMHMB52_A_PhLoad
Public PHIo As New MxMHMB52_A_PHIo
Public Ph7Tb As New MxMHMB52_A_Ph7Tb
Public PhTmpn As New MxMHMB52_A_PhTmpn
Public PhTmp5 As New MxMHMB52_A_PhTmp5
Public TbPh As New MxMHMB52_A_TbPh
Public SHldIo As New MxMHMB52_A_SHldIo
Public SHldFmtDteTit15Mth As New MxMHMB52_A_SHldFmtDteTit15Mth
Public SHldRpt As New MxMHMB52_A_SHldRpt
Public SHldSamp As New MxMHMB52_A_SHldSamp
Public SHldWsn As New MxMHMB52_A_SHldWsn
Public SHldStkDaysCalc As New MxMHMB52_A_SHldStkDaysCalc
Public SHldTmp7 As New MxMHMB52_A_SHldTmp7
Public SHldTarMths As New MxMHMB52_A_SHldTarMths
Public IsLoad As MxMHMB52_A_IsLoad
Public MB52Pm As New MxMHMB52_A_MB52Pm
Public SkuPm As New MxMHMB52_A_SkuPm
Public SkuIo As New MxMHMB52_A_SkuIO
Public ZHT0Pm As New MxMHMB52_A_ZHT0Pm
Public ZHT0Io As New MxMHMB52_A_ZHT0Io
Public ZHT0Load As New MxMHMB52_A_ZHT0Load
Public FcPm As New MxMHMB52_A_FcPm
Sub Main()
ClsWinAll
ClsFrmAll
DoCmd.OpenForm "Switchboard"
End Sub
Sub OpnFrmLoadSku():       OpnFrm "LoadSku":       End Sub
Sub OpnFrmLoadPh():        OpnFrm "LoadPh":        End Sub
Sub OpnFrmLoadZHT0():      OpnFrm "LoadZHT0":      End Sub
Sub OpnFrmLoadFc():        OpnFrm "LoadFc":        End Sub
Sub OpnFrmMacauCalcRate(): OpnFrm "MacauCalcRate": End Sub
Sub OpnFrmRpt():           OpnFrm "Rpt":           End Sub
Sub SkuRpt():
Dim X As New MxMHMB52_A_SkuRpt
X.SkuRpt
End Sub
