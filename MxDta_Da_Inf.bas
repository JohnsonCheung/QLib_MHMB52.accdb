Attribute VB_Name = "MxDta_Da_Inf"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Inf."
Public Const vbFldSep$ = ""

Function ValFeq(D As Drs, FldnSel$, F$, Eq)
Dim CixF%: CixF = CixDrs(D, F)
Dim DrSel: DrSel = WDrSel(D.Dy, CixF, Eq): If IsEmpty(DrSel) Then Exit Function
Dim CixSel%: CixSel = IxEle(D.Fny, FldnSel)
ValFeq = DrSel(CixSel) '<=
End Function
Private Function WDrSel&(Dy(), Cix%, Eq)
Dim Dr, RixSel&: For Each Dr In Itr(Dy)
    If Dr = Eq Then
        WDrSel = Dr
        Exit Function
    End If
    RixSel = RixSel + 1
Next
End Function

Function WdtDcDrs%(A As Drs, C$): WdtDcDrs = AyWdt(DcStrDrs(A, C)): End Function
Function HasRecDy2V(Dy(), C1, C2, V1, V2) As Boolean
Dim Dr
For Each Dr In Itr(Dy)
    If Dr(C1) = V1 Then
        If Dr(C2) = V2 Then
            HasRecDy2V = True
            Exit Function
        End If
    End If
Next
End Function

Function IsSamNCol(A As Drs, NCol%) As Boolean
Dim Dr
For Each Dr In Itr(A.Dy)
    If Si(Dr) = NCol Then Exit Function
Next
IsSamNCol = True
End Function

Function ResiDrs(A As Drs, NCol%) As Drs
If IsSamNCol(A, NCol) Then ResiDrs = A: Exit Function
Dim O As Drs, U%, Dr, J%
U = NCol - 1
For J = 0 To UB(O.Dy)
    Dr = O.Dy(J)
    ReDim Preserve Dr(U)
    O.Dy(J) = Dr
Next
End Function

Function DcLngDrsEq(D As Drs, C$, V, FldnSel$) As Long()
DcLngDrsEq = DcLngDrs(DwEqSelFf(D, C, V, FldnSel), FldnSel)
End Function
