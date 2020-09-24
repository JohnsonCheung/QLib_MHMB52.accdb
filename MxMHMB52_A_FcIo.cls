VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_FcIo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_FcIo."
Private Fun As New MxMHMB52_MHFun
Property Get WsnMH$(): WsnMH = "Table": End Property
Function Pthi$()
Static P$
If P = "" Then
    P = "C:\Users\Public\Logistic\StockHolding8\SAPDownloadExcel\Forecast\"
    'P = "C:\Users\Public\Logistic\SAPData\Forecast\"
End If
Pthi = P
End Function
Function FnFxi$(A As StmYm)
With A
FnFxi = FmtQQ("Fc ? ?.xlsx", Hypym(.Y, .M), Fun.Stm2Stm(.Stm))
End With
End Function
Function Fxi$(A As StmYm):    Fxi = Pthi & FnFxi(A):                    End Function
Function Ptho$():            Ptho = PthOup:                             End Function
Function FxoLas$():        FxoLas = Fxo(MH.FcLas.Ym):                   End Function
Function Fxo$(A As Ym):       Fxo = Ptho & FnFxo(A):                    End Function
Function FnFxo$(A As Ym):   FnFxo = "Forecast " & HypymYm(A) & ".xlsx": End Function

Function Fxw$(A As StmYm): Fxw = Pthw & "Wrk " & FnFxi(A): End Function
Function Pthw$()
Static P$
If P = "" Then
    P = Pthi & "Wrk\"
    If Dir(P, vbDirectory) = "" Then MkDir P
End If
Pthw = P
End Function
Function Wsnn$(Stm$)
If Stm = "M" Then
    Wsnn = WsnMH
Else
    Wsnn = WsnUD(86) & " " & WsnUD(87)
End If
End Function
Function WsnUD$(Co As Byte)
Const CSub$ = CMod & "WsnUD"
Select Case Co
Case 87: WsnUD = "MhdMO - Std Case"
Case 86: WsnUD = "MhdHK - Std Case"
Case Else: Thw CSub, "Co must be 86 or 87", "Co", Co
End Select
End Function

Function FxyInpMH() As String(): FxyInpMH = Ffny(Pthi, "Fc 20??-?? MH.xlsx"): End Function
Function FxyInpUD() As String(): FxyInpUD = Ffny(Pthi, "Fc 20??-?? UD.xlsx"): End Function
Function FxyInp() As String():     FxyInp = SyAdd(FxyInpUD, FxyInpMH):        End Function
