Attribute VB_Name = "MxXls_Rg_RgOp_Bdr"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Rg_Bdr."
Sub BdrBottom(A As Range)
BdrRg A, xlEdgeBottom
If A.Row < MaxRno Then
    BdrRg RgR(A, NRowRg(A) + 1), xlEdgeTop
End If
End Sub

Sub BdrInside(A As Range)
BdrRg A, xlInsideHorizontal
BdrRg A, xlInsideVertical
End Sub

Sub BdrLeft(A As Range)
BdrRg A, xlEdgeLeft
If A.Column > 1 Then
    BdrRg RgC(A, 0), xlEdgeRight
End If
End Sub

Sub BdrRight(A As Range)
BdrRg A, xlEdgeRight
If A.Column < MaxCno Then
    BdrRg RgC(A, NColRg(A) + 1), xlEdgeLeft
End If
End Sub

Sub BdrTop(A As Range)
BdrRg A, xlEdgeTop
If A.Row > 1 Then
    BdrRg RgR(A, 0), xlEdgeBottom
End If
End Sub

Sub BdrRg(A As Range, Ix As XlBordersIndex, Optional Wgt As XlBorderWeight = xlMedium)
With A.Borders(Ix)
  .LineStyle = xlContinuous
  .Weight = Wgt
End With
End Sub
Private Sub B_BdrAround()
GoSub ZZ
Exit Sub
ZZ:
    Dim S As Worksheet: Set S = WsNw
    Maxv S.Application
    BdrAround S.Range("B2:C3")
    Return
End Sub
Sub BdrNone(A As Range, Ix As XlBordersIndex): A.Borders(Ix).LineStyle = xlLineStyleNone: End Sub
Sub BdrAround(A As Range)
BdrLeft A
BdrRight A
BdrTop A
BdrBottom A
End Sub

Sub BdrClr(A As Range)
BdrNone A, xlInsideHorizontal
BdrNone A, xlInsideVertical
BdrNone A, xlEdgeLeft
BdrNone A, xlEdgeRight
BdrNone A, xlEdgeBottom
BdrNone A, xlEdgeTop
End Sub

Sub BdrRgLeft(R As Range):  BdrRg R, xlEdgeLeft:  End Sub
Sub BdrRgRight(R As Range): BdrRg R, xlEdgeRight: End Sub

Sub BdrRgAy(A() As Range, Ix As XlBordersIndex, Optional Wgt As XlBorderWeight = xlMedium)
Dim I: For Each I In Itr(A)
    BdrRg CvRg(I), Ix, Wgt
Next
End Sub

Function BdrLoAround(A As ListObject)
Dim R As Range
Set R = RgMoreTop(A.DataBodyRange)
If A.ShowTotals Then Set R = RgMoreBelow(R)
BdrAround R
End Function
