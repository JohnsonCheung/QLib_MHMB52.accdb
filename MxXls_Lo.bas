Attribute VB_Name = "MxXls_Lo"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Lo."

Function LoWs(A As Worksheet, Lon$) As ListObject: Set LoWs = ItoFstNm(A.ListObjects, Lon): End Function
Function LoWb(A As Workbook, Lon$) As ListObject
Dim S As Worksheet: For Each S In A.Sheets
    Set LoWb = LoWs(S, Lon)
    If Not IsNothing(LoWb) Then Exit Function
Next
End Function
Function LoFst(B As Worksheet) As ListObject: Set LoFst = ItvFst(B.ListObjects): End Function 'Return LoOpt
