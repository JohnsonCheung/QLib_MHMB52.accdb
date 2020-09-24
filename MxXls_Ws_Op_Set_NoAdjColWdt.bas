Attribute VB_Name = "MxXls_Ws_Op_Set_NoAdjColWdt"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Ws_Op_Set_NoAdjColWdt."
Sub SetColWdtNoAdjWb(B As Workbook)
Dim S As Worksheet: For Each S In B.Sheets
    SetColWdtNoAdjWs S
Next
End Sub
Sub SetColWdtNoAdjWs(S As Worksheet)
Dim L As ListObject: For Each L In S.ListObjects
    SetColWdtNoAdjLo L
Next
End Sub
Sub SetColWdtNoAdjLo(L As ListObject): L.QueryTable.AdjustColumnWidth = False: End Sub
