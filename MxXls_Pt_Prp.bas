Attribute VB_Name = "MxXls_Pt_Prp"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Pt_Prp."

Function FstPt(Ws As Worksheet) As PivotTable
Set FstPt = Ws.PivotTables(1)
End Function
Function CvPt(A) As PivotTable: Set CvPt = A: End Function
Sub FmtPivDfFmt(Pt As PivotTable, FfDta$, Fmt$)
Dim F: For Each F In Tmy(FfDta)
    SetPivDfFmt Pt, F, Fmt
Next
End Sub
Sub SetPivDfFmt(Pt As PivotTable, F, Fmt$): Pt.DataFields(F).NumberFormat = Fmt: End Sub
Function PtDtaFny(Pt As PivotTable) As String()
Dim F As PivotField: For Each F In Pt.DataFields
    PushS PtDtaFny, F.Name
Next
End Function
Sub FmtPivDfWdt(Pt As PivotTable, FF$, W)
Dim C: For Each C In FnyFF(FF)
    SetPivDfWdt Pt, C, W
Next
End Sub
Sub SetPivDfWdt(Pt As PivotTable, F, W):                                    RgEntColzPtFld(Pt, F).ColumnWidth = W:   End Sub
Function RgEntColzPtFld(Pt As PivotTable, F) As Range: Set RgEntColzPtFld = Pt.DataFields(F).DataRange.EntireColumn: End Function
Sub FmtPivRfWdt(Pt As PivotTable, FfRow$, W)
Dim C: For Each C In Tmy(FfRow)
    SetPivRfWdt Pt, C, W
Next
End Sub
Sub SetPivRfWdt(Pt As PivotTable, Rf, W): EntPivRfCol(Pt, Rf).ColumnWidth = W: End Sub
Function EntPivRfCol(Pt As PivotTable, PivRfn) As Range
Dim F As PivotField: Set F = Pt.PivotFields(PivRfn)
If F.Orientation <> xlRowField Then ThwPm "EntPivRfCol", "PivRfhn", PivRfn, "xlRowField", , "PivRfh.Orientation is not a row"
Set EntPivRfCol = F.DataRange.EntireColumn
End Function

Function PtFny(Pt As PivotTable) As String()
Dim F As PivotField: For Each F In Pt.PivotFields
    PushS PtFny, F.Name
Next
End Function

Sub ClrPt(Ws As Worksheet)
Dim Pt As PivotTable: For Each Pt In Ws.PivotTables
    Pt.TableRange2.ClearContents
Next
End Sub
