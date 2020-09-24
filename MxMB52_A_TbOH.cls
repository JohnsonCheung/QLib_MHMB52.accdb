VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMB52_A_TbOH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMB52_A_TbOH."

Function DteLas() As Date:          DteLas = DteYmd(YmdLas):                                  End Function
Function YmdLas() As Ymd:           YmdLas = YmdYymmdd(WYymmddLas):                           End Function
Private Function WYymmddLas&(): WYymmddLas = ValQC("Select Max(YY*10000+MM*100+DD) from OH"): End Function

Function WhLas$():                          WhLas = " where " & WBepr_Ymd(YmdLas):   End Function
Function WhYmd$(A As Ymd, Optional Alias$): WhYmd = " where " & WBepr_Ymd(A, Alias): End Function
Private Function WBepr_Ymd$(A As Ymd, Optional Alias$)
With A
WBepr_Ymd = WBepr_3Ymd(.Y, .M, .D, Alias)
End With
End Function
Private Function WBepr_3Ymd$(Y As Byte, M As Byte, D As Byte, Optional Alias$)
Const WTp$ = "{A}YY={Y} and {A}MM={M} and {A}DD={D}"
WBepr_3Ymd = FmtMacro(WTp, "A Y M D", AliasIf(Alias), Y, M, D)
End Function

Function WhCoYmd$(A As CoYmd): WhCoYmd = " where " & WBepr_Ymd(A.Ymd) & " and Co=" & A.Co: End Function
