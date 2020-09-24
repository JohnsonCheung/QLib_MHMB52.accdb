VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_SkuIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Function Fxo$()
Fxo = MH.MB52IO.Ptho & FmtQQ("OnHand SkuLis(?) Gen@?", _
HypymdDte(MH.TbOH.DteLas), _
Format(Now, "YYYY-MM-DD HHMM") & ".xlsx")
End Function
Function Fxi$(): Fxi = MH.SkuPm.Fxi: End Function
