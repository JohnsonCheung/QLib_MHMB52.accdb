Attribute VB_Name = "MxXls_Rfh"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Rfh."
Sub RfhFx(Fx$, Fb$)
Dim B As Workbook: Set B = WbFx(Fx)
RfhWb B, Fb
DltFfnIf Fx
B.SaveAs Fx
B.Close
End Sub
Sub RfhWb(B As Workbook, Fb)
B.Application.DisplayAlerts = False
WRfhWc B, Fb
WRfhPc B
WRfhWs B
FmtLoWbStd B
ClsWcWb B
Dim L As ListObject
L.QueryTable.C
DltWc B
B.Application.DisplayAlerts = True
End Sub

Private Sub WRfhWc(B As Workbook, Fb)
Dim C As WorkbookConnection: For Each C In B.Connections
    SetWcFb C, Fb
    C.OLEDBConnection.BackgroundQuery = False
    C.OLEDBConnection.Refresh
Next
End Sub
Private Sub WRfhPc(B As Workbook)
Dim C As PivotCache: For Each C In B.PivotCaches
    C.MissingItemsLimit = xlMissingItemsNone
    C.Refresh
Next
End Sub
Private Sub WRfhWs(B As Workbook)
Dim S As Worksheet: For Each S In B.Sheets
    Dim Q As QueryTable: For Each Q In S.QueryTables
        Q.Refresh False
        If Q.QueryType = xlOLEDBQuery Then Q.MaintainConnection = False
    Next
    Dim P As PivotTable: For Each P In S.PivotTables: P.RefreshTable: Next
    Dim L As ListObject: For Each L In S.ListObjects: L.Refresh: Next
Next
End Sub
