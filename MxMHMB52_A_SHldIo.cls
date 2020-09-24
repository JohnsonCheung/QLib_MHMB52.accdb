VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_SHldIo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_SHldIo."

Private WPm As MxMHMB52_A_MB52Pm
Function Ptho$():                             Ptho = MH.MB52IO.Ptho:                                       End Function
Function Tp$():                                 Tp = PthTp & "Stock Holding Template.xlsx":                End Function
Function Fxo$(A As CoYmd):                     Fxo = Ptho & WFn(A.Co, A.Ymd):                              End Function
Function Fxo86$(A As Ymd):                   Fxo86 = Ptho & WFn(86, A):                                    End Function
Function Fxo87$(A As Ymd):                   Fxo87 = Ptho & WFn(87, A):                                    End Function
Function Fxo861$(A As Ymd):                 Fxo861 = WPm.PthCpy1 & WFn(86, A):                             End Function
Function Fxo862$(A As Ymd):                 Fxo862 = WPm.PthCpy2 & WFn(86, A):                             End Function
Function Fxo871$(A As Ymd):                 Fxo871 = WPm.PthCpy1 & WFn(87, A):                             End Function
Function Fxo872$(A As Ymd):                 Fxo872 = WPm.PthCpy2 & WFn(87, A):                             End Function
Private Function WFn(Co As Byte, A As Ymd):    WFn = FmtQQ("Stock Holding ?(?00).xlsx", HypymdYmd(A), Co): End Function

Private Sub Class_Initialize()
Set WPm = MH.MB52Pm
End Sub
