Attribute VB_Name = "MxXls_Rg_Cell"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Rg_IsCell."
Function IsCellInRg(A As Range, Rg As Range) As Boolean
Dim R&, C%, R1&, R2&, C1%, C2%
R = A.Row
R1 = Rg.Row
If R < R1 Then Exit Function
R2 = R1 + NRowRg(Rg)
If R > R2 Then Exit Function
C = A.Column
C1 = Rg.Column
If C < C1 Then Exit Function
C2 = C1 + NColRg(Rg)
If C > C2 Then Exit Function
IsCellInRg = True
End Function

Function IsCellInRgAp(Cell As Range, ParamArray RgAp()) As Boolean
Dim Av(): Av = RgAp
'IsCellInRgAp = IsCellInRgAv(A, Av)
End Function

Function IsCellInRgAv(A As Range, RgAv()) As Boolean
Dim V
For Each V In RgAv
    If IsCellInRg(A, CvRg(V)) Then IsCellInRgAv = True: Exit Function
Next
End Function

Function CellBelow(Cell As Range, Optional N = 1) As Range:     Set CellBelow = RgRC(Cell, 1 + N, 1):     End Function
Function CellAbove(Cell As Range, Optional Above = 1) As Range: Set CellAbove = RgRC(Cell, 1 - Above, 1): End Function
Function CellRight(A As Range, Optional Right = 1) As Range:    Set CellRight = RgRC(A, 1, 1 + Right):    End Function


Private Sub B_SetCmt()
Dim R As Range: Set R = A1Nw
SetCmt R, "lskdfjsdlfk"
SetCmt R, "sdf"
End Sub
Sub SetCmt(R As Range, Cmt$)
If Not HasCmt(R) Then
    R.AddComment.Text Cmt
    Exit Sub
End If
Dim C As Comment: Set C = R.Comment
If C.Text = Cmt Then Exit Sub
C.Text Cmt, , True
End Sub
Function CmtRg$(R As Range)
On Error Resume Next
CmtRg = R.Comment
End Function
Function HasCmt(R As Range) As Boolean: HasCmt = Not IsNothing(R.Comment): End Function
Sub DltCmt(R As Range)
If Not IsNothing(R.Comment) Then R.Comment.Delete
End Sub
