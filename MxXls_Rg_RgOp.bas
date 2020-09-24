Attribute VB_Name = "MxXls_Rg_RgOp"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Rg_Prp."
Function AdrRg$(R As Range):                     AdrRg = R.Address(External:=True): End Function
Function IsRgSngRow(R As Range) As Boolean: IsRgSngRow = NRowRg(R) = 1: End Function
Function IsRgSngCol(R As Range) As Boolean: IsRgSngCol = NColRg(R) = 1: End Function
Function SqRg(R As Range) As Variant()
If NColRg(R) = 1 Then
    If NRowRg(R) = 1 Then
        Dim O()
        ReDim O(1 To 1, 1 To 1)
        O(1, 1) = R.Value
        SqRg = O
        Exit Function
    End If
End If
SqRg = R.Value
End Function

Function IsA1(R As Range) As Boolean
If R.Row <> 1 Then Exit Function
If R.Column <> 1 Then Exit Function
IsA1 = True
End Function

Function AdrWs$(R As Range): AdrWs = "'" & WsRg(R).Name & "'!" & R.Address: End Function ' AdrWs of @R.  AdrWs always with Wsn

Function TRcRg(R As Range) As TRc
With TRcRg
.R = R.Row
.C = R.Column
End With
End Function

Function RrccRg(R As Range) As Rrcc
With RrccRg
.R1 = R.Row
.R2 = .R1 + NRowRg(R) - 1
.C1 = R.Column
.C2 = .C1 + NColRg(R) - 1
End With
End Function

Function DrRg(Rg As Range, Optional R = 1) As Variant(): DrRg = DrSq(SqRg(RgR(Rg, R))): End Function

Function AdrA1$(R As Range):   AdrA1 = A1Rg(R).Address(External:=True): End Function
Function NRowRg&(R As Range): NRowRg = R.Rows.Count:    End Function
Function NColRg&(R As Range): NColRg = R.Columns.Count: End Function

Function ErCellVal$(S As Worksheet, Adr$, StrEpt$)
If S.Range(Adr).Value <> StrEpt Then ErCellVal = "Cell[" & Adr & "] should be [" & StrEpt & "] but now[" & S.Range(Adr).Value & "]"
End Function

Sub ChkIsCell(R As Range, Fun$): ThwFalse IsCell(R), Fun, "Given R is not a cell", "Rg-Address", R.Address: End Sub
Function IsCell(R As Range) As Boolean
If NRowRg(R) > 1 Then Exit Function
If NColRg(R) > 1 Then Exit Function
IsCell = True
End Function
