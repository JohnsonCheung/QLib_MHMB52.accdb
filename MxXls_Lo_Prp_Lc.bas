Attribute VB_Name = "MxXls_Lo_Prp_Lc"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Lo_Prp_Lc."

Function CellTit(L As ListObject, C) As Range: Set CellTit = A1Rg(CellAbove(L.ListColumns(C).Range)): End Function
