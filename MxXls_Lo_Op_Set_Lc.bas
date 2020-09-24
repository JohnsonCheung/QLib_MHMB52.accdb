Attribute VB_Name = "MxXls_Lo_Op_Set_Lc"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Lo_Op_Set_Lc."

Property Let Lcc_Wdt(L As ListObject, CC$, W):              Dim C: For Each C In Tmy(CC): Lc_Wdt(L, C) = W:  Next:   End Property
Sub SetLccFmt(L As ListObject, CC$, Fmt$):                  Dim C: For Each C In Tmy(CC): SetLcFmt L, C, Fmt: Next:  End Sub
Sub SetLccAsSum(L As ListObject, CC$):                      Dim C: For Each C In Tmy(CC): SetLcAsSum L, C:     Next: End Sub
Sub SetLccAsAvg(L As ListObject, CC$):                      Dim C: For Each C In Tmy(CC): SetLcAsAvg L, C:     Next: End Sub
Sub SetLccLvl(L As ListObject, CC$, Lvl As Byte):           Dim C: For Each C In Tmy(CC): SetLcLvl L, C, Lvl: Next:  End Sub
Sub SetLcAli(L As ListObject, C, A As eLofali):             RgDtaLc(L, C).HorizontalAliment = WCvXlHAli(A):          End Sub
Sub SetLcBdr(L As ListObject, C, A As eLofBdr):             BdrRg RgLc(L, C), WCvXlBdrIx(A):                         End Sub
Sub SetLcFmt(L As ListObject, C, Fmt$):                     RgDtaLc(L, C).NumberFormat = Fmt:                        End Sub
Sub SetLcLvl(L As ListObject, C, Optional Lvl As Byte = 2): RgLcEnt(L, C).OutlineLevel = Lvl:                        End Sub
Property Let Lc_Wdt(L As ListObject, C, W):                 RgLcEnt(L, C).ColumnWidth = W:                           End Property
Sub SetLcWdt(L As ListObject, Wdtln$)
Dim Ln$: Ln = Wdtln
Dim W%: W = ShfTm(Ln)
Lcc_Wdt(L, Ln) = W
End Sub
Sub SetLcWrp(L As ListObject, C, Wrp As Boolean): RgDtaLc(L, C).WrapText = Wrp:                 End Sub
Sub SetLcTot(L As ListObject, C, A As eLofagr):   Lc(L, C).TotalsCalculation = WCvXlTotCalc(A): End Sub
Sub SetLcCor(L As ListObject, C, Colr&):          RgDtaLc(L, C).Interior.Color = Colr:          End Sub
Sub SetLcSum(L As ListObject, SumFld$, FmFld$, ToFld$)
Dim Fml$
Stop
RgDtaLc(L, SumFld).Formula = Fml
End Sub
Sub SetLcLbl(L As ListObject, C, Lbl$)
Dim CellLbl As Range
Stop
CellLbl.Value = Lbl
End Sub
Sub SetLcAsSum(L As ListObject, C): X_SetCalc L, C, xlTotalsCalculationSum:     End Sub
Sub SetLcAsCnt(L As ListObject, C): X_SetCalc L, C, xlTotalsCalculationCount:   End Sub
Sub SetLcAsAvg(L As ListObject, C): X_SetCalc L, C, xlTotalsCalculationAverage: End Sub
Private Sub X_SetCalc(L As ListObject, C, Calc As XlTotalsCalculation)
L.ShowTotals = True: Lc(L, C).TotalsCalculation = Calc
End Sub
Private Function WCvXlBdrIx(A As eLofBdr) As XlBordersIndex
Stop
End Function
Private Function WCvXlHAli(A As eLofali) As XlHAlign
Stop
Const CSub$ = CMod & "WCvXlHAli"
Select Case A
Case "Left": WCvXlHAli = xlHAlignLeft
Case "Right": WCvXlHAli = xlHAlignRight
Case "Center": WCvXlHAli = xlHAlignCenter
'Case Else: Inf CSub, "Invalid Ali", "Valid Ali", Lofaliss: Exit Function
End Select
End Function
Private Function WCvXlTotCalc(A As eLofagr) As XlTotalsCalculation
Const CSub$ = CMod & "WCvXlTotCalc"
'Fm SACnt : "Sum | Avg | Cnt" @@
Dim O As XlTotalsCalculation
Select Case A
Case "Sum": O = xlTotalsCalculationSum
Case "Avg": O = xlTotalsCalculationAverage
Case "Cnt": O = xlTotalsCalculationCount
Case Else: Inf CSub, "Invalid TotCalcStr", "TotCalcStr Valid-TotCalcStr", A, "Sum Avg Cnt": Exit Function
End Select
WCvXlTotCalc = O
End Function

Private Sub B_FmllnyLo()
Dim S As Worksheet: Set S = samp_mhmb52rptdta_Ws
Stop
D FmllnyLo(LoFst(S))
QuitWs S
End Sub

Private Sub B_FmlFny()
Dim S As Worksheet: Set S = samp_mhmb52rptdta_Ws
Dim L As Object: Set L = LoFst(S)
D FnyFml(L)
QuitWs S
End Sub
Function FnyFml(L As ListObject) As String() ' Return the Fny with formula according to the fst ListRow
If L.ListRows.Count = 0 Then Exit Function
Dim R As Range, J%, Fml$: For Each R In L.ListRows(1).Range
    J = J + 1
    Fml = R.Formula
    If ChrFst(Fml) = "=" Then PushI FnyFml, L.ListColumns(J).Name
Next
End Function

Private Sub B_FmllnLo()
Dim Wb As Workbook: Set Wb = samp_mhmb52rpt_Wb
Dim Ws As Worksheet: Set Ws = Wb.Worksheets("MacauBchRat")
Dim L As ListObject: Set L = Ws.ListObjects(1)
SetLoFmlln L, "Litre=[@Btl] * [@Size] / 100"
SetLoFmlln L, "LitreHKD=[@Litre] * [@[MOP/Litre]] * [@[HKD/MOP]]"
SetLoFmlln L, "10%A=[@[XXX/Btl]] * [@Btl] * 0.1"
SetLoFmlln L, "10%B=[@Val] * 0.1"
SetLoFmlln L, "10%HKD=IF(ISBLANK([@[XXX/Btl]]),[@[10%B]],[@[10%A]])"
SetLoFmlln L, "HKD=[@LitreHKD] + [@[10%HKD]]"
SetLoFmlln L, "HKD/Ac=[@[Btl/Ac]]"
End Sub

Function FmllnyLo(L As ListObject) As String()
Dim F: For Each F In Itr(FnyFml(L))
    Dim R As Range: Set R = RgRC(L.ListColumns(F).Range, 2, 1)
    Dim Fml$: Fml = R.Formula
    PushI FmllnyLo, F & Fml
Next
End Function
Sub SetLoFmllny(L As ListObject, Fmllny$())
Dim Fmlln: For Each Fmlln In Itr(Fmllny)
    SetLoFmlln L, Fmlln
Next
End Sub
Sub SetLoFmlln(L As ListObject, Fmlln)
With BrkEq(Fmlln)
    SetLcFml L, .S1, "=" & .S2
End With
End Sub
Sub SetLcFml(L As ListObject, C$, Fml$)
If C = "" Then Exit Sub
If Not HasLc(L, C) Then Inf CSub, "No such Lc in Lo", "Lc Lon", C, L.Name: Exit Sub
L.ListColumns(C).DataBodyRange.Formula = Fml
End Sub

Sub InsRowLo(L As ListObject, Optional N = 1)
If L.ListRows.Count = 0 Then
    Stop
End If
Dim R As Range: Set R = RgRR(L.ListRows(1).Range, 1, N)
R.Insert xlShiftDown, xlFormatFromRightOrBelow
End Sub
Function RnyLc(C As ListColumn, DcSub) As Long()
Const CSub$ = CMod & "RnyLc"
Dim Dc$(): Dc = DcStrLc(C)
Dim J%: For J = 0 To UB(Dc)
    If HasEle(DcSub, Dc(J)) Then
        PushI RnyLc, J + 1
    End If
Next
End Function
