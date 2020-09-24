Attribute VB_Name = "MxVb_Dta_SPart"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_SPart."
Type SPart: Ix As Long: Sy() As String: End Type ' Deriving(Ctor Ay Opt) SPart represent a part of a bigger Sy
Type SPartOpt: Som As Boolean: SPart As SPart: End Type
Function SPart(Ix, Sy$()) As SPart
With SPart
    .Ix = Ix
    .Sy = Sy
End With
End Function
Function AddSPart(A As SPart, B As SPart) As SPart(): PushSPart AddSPart, A: PushSPart AddSPart, B: End Function
Sub PushSPartAy(O() As SPart, A() As SPart): Dim J&: For J = 0 To SPartUB(A): PushSPart O, A(J): Next: End Sub
Sub PushSPart(O() As SPart, M As SPart): Dim N&: N = SPartSI(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SPartSI&(A() As SPart): On Error Resume Next: SPartSI = UBound(A) + 1: End Function
Function SPartUB&(A() As SPart): SPartUB = SPartSI(A) - 1: End Function
Function SPartOpt(Som, A As SPart) As SPartOpt: With SPartOpt: .Som = Som: .SPart = A: End With: End Function
Function SomSPart(A As SPart) As SPartOpt: SomSPart.Som = True: SomSPart.SPart = A: End Function
