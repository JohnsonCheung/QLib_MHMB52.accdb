VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_FcWh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_FcWh."

Function WhStmYm$(A As StmYm):        WhStmYm = MH.FcWh.WhYm(A.Y, A.M) & QpAndFeq("Stm", A.Stm): End Function
Function WhYm$(Y As Byte, M As Byte):    WhYm = Wh("VerYY=" & Y & " and VerMM=" & M):            End Function
Function WhCoYm$(A As CoYm):           WhCoYm = WhYm(A.Y, A.M) & QpAndFeq("Co", A.Co):           End Function
