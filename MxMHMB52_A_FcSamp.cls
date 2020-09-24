VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_FcSamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_FcSamp."


Function FbOup$():                FbOup = sampFfn("SampFcOup.accdb"): End Function ' #Fc-Oup-database-file# a database file that contains @* to produce FcRpt
Function DbOup() As Database: Set DbOup = Db(FbOup):                  End Function ' #Fc-Oup-database# a database that contains @* to produce FcRpt

Function FxRpt$():                FxRpt = sampFfn("SampFc.xlsx"): End Function
Function WbRpt() As Workbook: Set WbRpt = WbFx(FxRpt):            End Function

Function LoSku() As ListObject: Set LoSku = LoWb(WbRpt, "Lo_FcSku"): End Function
