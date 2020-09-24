Attribute VB_Name = "MxXls_Rg_RgPrp_RgOffset"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Rg_Offset."

Function C2Rg%(Rg As Range):                                         C2Rg = Rg.Column + Rg.Columns.Count - 1:                   End Function
Function RgCEnt(Rg As Range, C) As Range:                      Set RgCEnt = RgC(Rg, C).EntireColumn:                            End Function
Function RgCCEnt(Rg As Range, C1, C2) As Range:               Set RgCCEnt = RgCC(Rg, C1, C2).EntireColumn:                      End Function
Function RgREnt(Rg As Range, Optional R = 1) As Range:         Set RgREnt = RgR(Rg, R).EntireRow:                               End Function
Function RgRREnt(Rg As Range, R1, R2) As Range:               Set RgRREnt = RgRR(Rg, R1, R2).EntireRow:                         End Function
Function RgFstCol(Rg As Range) As Range:                     Set RgFstCol = RgC(Rg, 1):                                         End Function
Function RgFstRow(Rg As Range) As Range:                     Set RgFstRow = RgR(Rg, 1):                                         End Function
Function RccRg(Rg As Range) As RCC:                                 RccRg = RCC(Rg.Row, Rg.Column, C2Rg(Rg)):                   End Function
Function RgC(Rg As Range, Optional C = 1) As Range:               Set RgC = RgCC(Rg, C, C):                                     End Function
Function RgCC(Rg As Range, C1, C2) As Range:                     Set RgCC = RgRCRC(Rg, 1, C1, NRowRg(Rg), C2):                  End Function
Function RgCRR(Rg As Range, C, R1, R2) As Range:                Set RgCRR = RgRCRC(Rg, R1, C, R2, C):                           End Function
Function RgMoreBelow(Rg As Range, Optional N% = 1):       Set RgMoreBelow = RgRR(Rg, 1, NRowRg(Rg) + N):                        End Function
Function RgR(Rg As Range, Optional R = 1) As Range:               Set RgR = RgRR(Rg, R, R):                                     End Function
Function RgRC(Rg As Range, R, C) As Range:                       Set RgRC = Rg.Cells(R, C):                                     End Function
Function RgRCC(Rg As Range, R, C1, C2) As Range:                Set RgRCC = RgRCRC(Rg, R, C1, R, C2):                           End Function
Function RgRCRC(Rg As Range, R1, C1, R2, C2) As Range:         Set RgRCRC = WsRg(Rg).Range(RgRC(Rg, R1, C1), RgRC(Rg, R2, C2)): End Function
Function RgRR(Rg As Range, R1, R2) As Range:                     Set RgRR = RgRCRC(Rg, R1, 1, R2, NColRg(Rg)):                  End Function
Function RgLessTop(Rg As Range, Optional N = 1) As Range:   Set RgLessTop = RgMoreTop(Rg, -N):                                  End Function
Function RgMoreTop(Rg As Range, Optional N = 1) As Range
Dim O As Range
Set O = RgRR(Rg, 1 - N, NRowRg(Rg))
Set RgMoreTop = O
End Function
Function ValAt(At As Range): ValAt = A1Rg(At).Value: End Function
Function RgAtDtaDown(At As Range) As Range
If IsEmpty(ValAt(At)) Then
    Set RgAtDtaDown = A1Rg(At)
Else
    Dim R2: R2 = 1
    Set RgAtDtaDown = RgRCRC(At, 1, 1, R2, 1)
End If
End Function
Function RgAtDtaRight(At As Range) As Range

End Function
Function CnoRgFstDta%(Rg As Range): CnoRgFstDta = CnoWsFstDta(WsRg(Rg), RccRg(Rg)): End Function
Function CnoWsFstDta%(S As Worksheet, RCC As RCC)
With RCC
Dim Cno%: For Cno = .C1 To .C2
    If Not IsEmpty(RgWsRC(S, .R, Cno).Value) Then CnoWsFstDta = Cno: Exit Function
Next
End With
End Function

Function RgHoriAt(At As Range) As Range

End Function
