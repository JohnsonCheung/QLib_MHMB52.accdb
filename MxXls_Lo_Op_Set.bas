Attribute VB_Name = "MxXls_Lo_Op_Set"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Lo_Op_Set."
Type ColNDecDig: Coln As String: NDecDig As Byte: End Type ' Deriving(Ay)
Type ColAgr: CntColn As String: SumColy() As ColNDecDig: End Type
Private Sub PushColNDecDig(O() As ColNDecDig, M As ColNDecDig)
Dim N&: N = ColNDecDigSI(O)
ReDim Preserve O(N)
O(N) = M
End Sub
Private Function ColNDecDigUB&(A() As ColNDecDig): ColNDecDigUB = ColNDecDigSI(A) - 1: End Function
Private Function ColNDecDigSI&(A() As ColNDecDig)
On Error Resume Next
ColNDecDigSI = UBound(A) + 1
End Function

Sub SetLoAutoFit(L As ListObject, Optional MaxW = 100)
Dim C As Range: Set C = RgLoAllColEnt(L)
C.AutoFit
Dim EntC As Range, J%
For J = 1 To C.Columns.Count
   Set EntC = RgCEnt(C, J)
   If EntC.ColumnWidth > MaxW Then EntC.ColumnWidth = MaxW
Next
End Sub
Sub SetLoTot(L As ListObject)
Dim C As ColAgr: C = ColAgr(L)
Dim J%: For J = 0 To ColNDecDigUB(C.SumColy)
    With C.SumColy(J)
    L.ListColumns(.Coln).TotalsCalculation = xlTotalsCalculationSum
    RgDtaLc(L, .Coln, InlTot:=True).NumberFormat = WNbrFmt(.NDecDig)
    End With
Next
If C.CntColn <> "" Then L.ListColumns(C.CntColn).TotalsCalculation = xlTotalsCalculationCount
L.ShowTotals = True
End Sub
Private Function WNbrFmt$(NDig As Byte)
If NDig = 0 Then
    WNbrFmt = "#,##0"
Else
    WNbrFmt = "#,##0." & Pad0(0, NDig)
End If
End Function
Private Function ColAgr(L As ListObject) As ColAgr
Dim C As ListColumn: For Each C In L.ListColumns
    Dim Dc(): Dc = DcLc(C)
    If IsAyNbr(Dc) Then
        Dim M As ColNDecDig
        M.Coln = C.Name
        M.NDecDig = WNDecDigNbry(Dc)
        PushColNDecDig ColAgr.SumColy, M
    Else
        If ColAgr.CntColn = "" Then ColAgr.CntColn = C.Name
    End If
Next
End Function
Private Function WNDecDigNbry(Nbry) As Byte
Dim O As Byte: Dim Nbr: For Each Nbr In Itr(Nbry)
    O = Max(O, WNDecDig(Nbr)): If O = 3 Then WNDecDigNbry = 3: Exit Function
Next
WNDecDigNbry = O
End Function
Private Function WNDecDig(Nbr) As Byte
If Not IsDbl(Nbr) Then Exit Function
Dim D#: D = Nbr - Fix(Nbr): If D = 0 Then Exit Function
WNDecDig = Min(Len(CStr(D)) - 2, 3)
End Function
Sub SetLon(L As ListObject, Lon$)
If Lon = "" Then Exit Sub
L.Name = Lon
End Sub
