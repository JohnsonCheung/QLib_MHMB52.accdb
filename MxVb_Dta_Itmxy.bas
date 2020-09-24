Attribute VB_Name = "MxVb_Dta_Itmxy"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Itmxy."
Type Itmxy: Itm As String: Ixy() As Long: End Type 'Deriving(Ctor Ay)

Function DisIxyItmxyAy(I() As Itmxy) As Long()
Dim J%: For J = 0 To ItmxyUB(I)
    PushNoDupIAy DisIxyItmxyAy, I(J).Ixy
Next
End Function
Function Itmxy(Itm, Ixy&()) As Itmxy
With Itmxy
    .Itm = Itm
    .Ixy = Ixy
End With
End Function
Function AddItmxy(A As Itmxy, B As Itmxy) As Itmxy(): PushItmxy AddItmxy, A: PushItmxy AddItmxy, B: End Function
Sub PushItmxyAy(O() As Itmxy, A() As Itmxy): Dim J&: For J = 0 To ItmxyUB(A): PushItmxy O, A(J): Next: End Sub
Sub PushItmxy(O() As Itmxy, M As Itmxy): Dim N&: N = ItmxySI(O): ReDim Preserve O(N): O(N) = M: End Sub
Function ItmxySI&(A() As Itmxy): On Error Resume Next: ItmxySI = UBound(A) + 1: End Function
Function ItmxyUB&(A() As Itmxy): ItmxyUB = ItmxySI(A) - 1: End Function
