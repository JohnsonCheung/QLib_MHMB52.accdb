Attribute VB_Name = "MxXls_Lo_Op_Clr"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Lo_Op_Clr."

Sub ClrLo(L As ListObject)
If Right3(L.Name) <> "Lo_" Then Exit Sub
Dim R1 As Range
Set R1 = L.DataBodyRange
If NRowRg(R1) >= 2 Then
    RgRR(R1, 2, NRowRg(R1)).EntireRow.Delete
End If
End Sub

Sub ClrLoWs(S As Worksheet)
Dim L As ListObject: For Each L In S.ListObjects
    ClrLo L
Next
End Sub

Sub ClrLoWb(B As Workbook)
Dim S As Worksheet: For Each S In B.Sheets
    ClrLoWs S
Next
End Sub

Private Sub B_ClrLoFx()
ClrLoFx MH.FcIO.FxoLas
End Sub
Sub ClrLoFx(Fx)
Dim O As Workbook: Set O = WbFx(Fx)
ClrLoWb O
SavWb O
ClsWbNoSav O
End Sub
