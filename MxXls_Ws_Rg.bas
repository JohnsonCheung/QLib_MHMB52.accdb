Attribute VB_Name = "MxXls_Ws_Rg"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Ws_Rg."
Function RgWsC(S As Worksheet, C) As Range:                     Set RgWsC = S.Columns(C).EntireColumn:                             End Function
Function RgWsCC(S As Worksheet, C1, C2) As Range:              Set RgWsCC = RgWsRCC(S, 1, C1, C2).EntireColumn:                    End Function
Function RgWsCRR(S As Worksheet, C, R1, R2) As Range:         Set RgWsCRR = RgWsRCRC(S, R1, C, R2, C):                             End Function
Function RgWsRC(S As Worksheet, R, C) As Range:                Set RgWsRC = S.Cells(R, C):                                         End Function
Function RgWsRCC(S As Worksheet, R, C1, C2) As Range:         Set RgWsRCC = RgWsRCRC(S, R, C1, R, C2):                             End Function
Function RgWsRCRC(S As Worksheet, R1, C1, R2, C2) As Range:  Set RgWsRCRC = S.Range(RgWsRC(S, R1, C1), RgWsRC(S, R2, C2)):         End Function
Function RgWsRR(S As Worksheet, R1, R2) As Range:              Set RgWsRR = S.Range(RgWsRC(S, R1, 1), RgWsRC(S, R2, 1)).EntireRow: End Function
Function RgWsDtaRg(S As Worksheet) As Range:                Set RgWsDtaRg = S.Range(A1Ws(S), LasCell(S)):                          End Function
