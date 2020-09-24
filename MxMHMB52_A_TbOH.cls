VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_TbOH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_TbOH."

Function DteLas() As Date: DteLas = DteYmd(YmdLas): End Function
Function YmdLas() As Ymd:
Dim YymmddLas&: YymmddLas = ValQC("Select Max(YY*10000+MM*100+DD) from OH")
YmdLas = YmdYymmdd(YymmddLas)
End Function

Function WhLas$(Optional Alias$):           WhLas = Wh(WBeprYmd(YmdLas, Alias)): End Function
Function WhYmd$(A As Ymd, Optional Alias$): WhYmd = Wh(WBeprYmd(A, Alias)):      End Function
Private Function WBeprYmd$(A As Ymd, Optional Alias$)
With A
WBeprYmd = WBprYmd3(.Y, .M, .D, Alias)
End With
End Function
Private Function WBprYmd3$(Y As Byte, M As Byte, D As Byte, Optional Alias$)
Const WTp$ = "{A}YY={Y} and {A}MM={M} and {A}DD={D}"
WBprYmd3 = FmtMacro(WTp, "A Y M D", AliasIf(Alias), Y, M, D)
End Function

Function WhCoYmd$(A As CoYmd): WhCoYmd = Wh(WBeprYmd(A.Ymd) & QpAndFeq("Co", A.Co)): End Function
