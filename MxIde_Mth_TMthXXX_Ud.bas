Attribute VB_Name = "MxIde_Mth_TMthXXX_Ud"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_TMthXXX_Ud."
Type TMth: Mthn As String: ShtTy As String: ShtMdy As String: End Type ' Deriving(Ay)
Type TMthmdn: TMth As TMth: Mdn As String: End Type 'Deriving(Ay)
Type TMthymdn: TMthy() As TMth: Mdn As String: End Type 'Deriving(Ay Ctor)
Function SiTMthmd&(A() As TMthmdn): On Error Resume Next: SiTMthmd = UBound(A) + 1: End Function
Function UbTMthmd&(A() As TMthmdn): UbTMthmd = SiTMthmd(A) - 1: End Function
Sub PushTMthMd(O() As TMthmdn, M As TMthmdn): Dim N&: N = SiTMthmd(O): ReDim Preserve O(N): O(N) = M: End Sub
Function TMthmdn(Mdn, TMth As TMth) As TMthmdn
With TMthmdn
    .Mdn = Mdn
    .TMth = TMth
End With
End Function
Function UbTMth&(A() As TMth): UbTMth = SiTMth(A) - 1: End Function
Function SiTMth&(A() As TMth): On Error Resume Next: SiTMth = UBound(A) + 1: End Function
Sub PushTMth(O() As TMth, M As TMth): Dim N&: N = SiTMth(O): ReDim Preserve O(N): O(N) = M: End Sub
Sub PushTMthy(O() As TMth, M() As TMth): Dim J&: For J = 0 To UbTMth(M): PushTMth O, M(J): Next: End Sub
Function TMth(Mthn, ShtTy, ShtMdy) As TMth
With TMth
    .Mthn = Mthn
    .ShtTy = ShtTy
    .ShtMdy = ShtMdy
End With
End Function
Function Mi2TyNmTMth$(M As TMth): Mi2TyNmTMth = M.ShtTy & " " & M.Mthn: End Function
