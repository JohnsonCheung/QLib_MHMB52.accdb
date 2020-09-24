Attribute VB_Name = "MxXls_Wb_Ixws"
Option Compare Text
Option Explicit
Private Sub B_EnsIxws()
Dim B As Workbook: Set B = WbNw
WsAdd B
WsWb(B, "Sheet2").Range("A1").Value = "This is sheet2"
EnsIxws B
Maxv B.Application
End Sub
Sub EnsIxws(B As Workbook, Optional Ixwsn$ = "Index") 'Create or Set IdxWs..
'Ixws is second Ws of a Wb (first is ReadMe)
'      it is A1 = Index
'      it is A2 = a list of worksheet name (With WsCdn like Ws*) in vertical with hyper link
'      Note the WsCdn is set by !CrtFxWc
Dim S As Worksheet
    Set S = WsWb(B, Ixwsn)
    If IsNothing(S) Then Set S = WsAdd(B, Ixwsn, ePosWsBef, "Readme")
    WEns S
End Sub
Private Sub WEns(S As Worksheet)
Dim N$(), At As Range, B As Workbook
   Set B = WbWs(S)
       N = AeEle(Wny(B), S.Name)
  Set At = S.Range("A2")
           PutAyVert N, At
           WAddHyp At, N
End Sub
Private Sub WAddHyp(At As Range, Wny$())
Dim B As Workbook: Set B = WbRg(At)
Dim RgIx As Range: Set RgIx = A1Ws(WsWb(B, WsRg(At).Name))
Dim Wsn, J%: For Each Wsn In Wny
    J = J + 1
    Dim RgAtIx As Range: Set RgAtIx = At(J, 1)
    HypLnkRgPr RgAtIx, A1Ws(B.Sheets(Wsn))
Next
End Sub
