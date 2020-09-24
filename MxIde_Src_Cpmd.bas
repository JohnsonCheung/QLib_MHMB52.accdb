Attribute VB_Name = "MxIde_Src_Cpmd"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_TMdnBAft."
Type Cpmdl: Mdn As String: Befl As String: Aftl As String: End Type 'Deriving(Ay Ctor)
Type Cpmd: Mdn As String: Bef() As String: Aft() As String: End Type 'Deriving(Ay Ctor)
Type Cpmdlopt: Som As Boolean: Cpmdl As Cpmdl: End Type 'Deriving(Ay Ctor)
Type Cpmdopt: Som As Boolean: Cpmd As Cpmd: End Type 'Deriving(Ay Ctor)
Function Cpmdl(Mdn, Befl, Aftl) As Cpmdl
With Cpmdl
    .Aftl = Aftl
    .Mdn = Mdn
    .Befl = Befl
End With
End Function
Function SomCpmdl(Cpmdl As Cpmdl) As Cpmdlopt
With SomCpmdl:
    .Som = True
    .Cpmdl = Cpmdl
End With
End Function
Sub PushCpmdl(O() As Cpmdl, M As Cpmdl): Dim N&: N = SiCpmdl(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiCpmdl&(A() As Cpmdl): On Error Resume Next: SiCpmdl = UBound(A) + 1: End Function
Function UbCpmdl&(A() As Cpmdl): UbCpmdl = SiCpmdl(A) - 1: End Function
Sub PushCpmdlOpt(O() As Cpmdl, M As Cpmdlopt)
If M.Som Then PushCpmdl O, M.Cpmdl
End Sub

Function Cpmd(Mdn, Bef, Aft) As Cpmd
With Cpmd
    .Aft = Aft
    .Mdn = Mdn
    .Bef = Bef
End With
End Function
Function SomCpmd(Cpmd As Cpmd) As Cpmdopt
With SomCpmd:
    .Som = True
    .Cpmd = Cpmd
End With
End Function
Sub PushCpmd(O() As Cpmd, M As Cpmd): Dim N&: N = SiCpmd(O): ReDim Preserve O(N): O(N) = M: End Sub
Sub PushCpmdy(O() As Cpmd, M() As Cpmd): Dim J&: For J = 0 To UbCpmd(M): PushCpmd O, M(J): Next: End Sub
Function SiCpmd&(A() As Cpmd): On Error Resume Next: SiCpmd = UBound(A) + 1: End Function
Function UbCpmd&(A() As Cpmd): UbCpmd = SiCpmd(A) - 1: End Function
Sub PushCpmdOpt(O() As Cpmd, M As Cpmdopt)
If M.Som Then PushCpmd O, M.Cpmd
End Sub
