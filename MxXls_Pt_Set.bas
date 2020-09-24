Attribute VB_Name = "MxXls_Pt_Set"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Pt_Set."

Sub SetPtWdt(A As PivotTable, Colss$, Wdt As Byte)
If Wdt <= 1 Then Stop
Dim C: For Each C In Itr(SySs(Colss))
    RgcEntPt(A, C).ColumnWidth = Wdt
Next
End Sub
Sub SetPtOutln(T As PivotTable, Colss$, Optional Lvl As Byte = 2)
If Lvl <= 1 Then Stop
Dim F, C As VBComponent: For Each C In Itr(SySs(Colss))
    RgcEntPt(T, F).OutlineLevel = Lvl
Next
End Sub

Sub SetPtRepeatLbl(T As PivotTable, Rowss$)
Dim F: For Each F In Itr(SySs(Rowss))
    PivFld(T, F).RepeatLabels = True
Next
End Sub

Function RgcEntPt(A As PivotTable, PivColNm) As Range:  Set RgcEntPt = RgR(PivFldCol(A, PivColNm).DataRange, 1).EntireColumn: End Function
Function PivColEnt(Pt As PivotTable, Coln) As Range:   Set PivColEnt = PivFldCol(Pt, Coln).EntireColumn:                      End Function

Private Sub B_PtLo()
Dim At As Range, Lo As ListObject
Stop 'Set Lo = SampLo
Maxv PtLo(Lo, At, "A B", "C D", "F", "E").Application
Stop
End Sub
Function PtnLon$(Lon$)
Stop '
End Function
Function PtLo(A As ListObject, At As Range, Rowss$, Dtass$, Optional Colss$, Optional Pagss$) As PivotTable
If WbLo(A).FullName <> WbRg(At).FullName Then Stop: Exit Function
Dim O As PivotTable
Set O = LoPc(A).CreatePivotTable(TableDestination:=At, TableName:=PtnLon(A.Name))
With O
    .ShowDrillIndicators = False
    .InGridDropZones = False
    .RowAxisLayout xlTabularRow
End With
O.NullString = ""
SetPivFldOri O, Rowss, xlRowField
SetPivFldOri O, Colss, xlColumnField
SetPivFldOri O, Pagss, xlPageField
SetPivFldOri O, Dtass, xlDataField
Set PtLo = O
End Function

Function WbPt(A As PivotTable) As Workbook
Set WbPt = WbWs(WsPt(A))
End Function

Function WsPt(A As PivotTable) As Worksheet
Set WsPt = A.Parent
End Function
