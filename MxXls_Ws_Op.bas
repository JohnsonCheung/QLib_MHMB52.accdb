Attribute VB_Name = "MxXls_Ws_Op"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Ws_Op."

Sub DltColFm(S As Worksheet, FmCol):  RgWsCC(S, FmCol, LasCno(S)).Delete:               End Sub
Sub DltRowFm(S As Worksheet, FmRow):  RgWsRR(S, FmRow, LasRno(S)).Delete:               End Sub
Sub HidColFm(S As Worksheet, FmCol):  RgWsCC(S, FmCol, MaxCno).Hidden = True:           End Sub
Sub HidRowFm(S As Worksheet, FmRow&): RgWsRR(S, FmRow, MaxRno).EntireRow.Hidden = True: End Sub
