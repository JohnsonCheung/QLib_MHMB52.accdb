Attribute VB_Name = "MxXls_Lo_PutDrs"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Lo_PutDrs."

Private Sub B_PutDrsLo()
GoSub Z
Dim D As Drs, L As ListObject
Exit Sub

Z:
    D = sampDrs1
    Stop '    Set L = LoNwDrs(D, A1Nw)
    PutDrsLo D, L
    Maxv L.Application
End Sub

Sub PutDrsLo(D As Drs, L As ListObject)
ClrLo L
Dim Dy(): Dy = DrsSelFny(D, FnyLo(L)).Dy
RgDy Dy, RgRC(L.Range, 2, 1)
End Sub
