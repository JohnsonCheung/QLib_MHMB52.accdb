VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_SHldSamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_SHldSamp."

Function samp_shld_rpt_Fx$():                samp_shld_rpt_Fx = sampFfn("SampSHld.Xlsx"): End Function
Function samp_shld_rpt_Wb() As Workbook: Set samp_shld_rpt_Wb = WbFx(samp_shld_rpt_Fx):   End Function
