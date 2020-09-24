Attribute VB_Name = "MxDta_Da_Op_SrtKey"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Op_SrtKey."
Type Srkey: Ci As Integer: IsDes As Boolean: End Type 'Deriving(Ctor Ay)
Function Srkeyy(SSSCxiHyp$) As Srkey()
Dim CiyHyp$():
End Function
Function Srkey(Ci, IsDes) As Srkey
With Srkey
    .Ci = Ci
    .IsDes = IsDes
End With
End Function
Sub PushSrkeyy(O() As Srkey, M() As Srkey): Dim J&: For J = 0 To UbSrkey(M): PushSrkey O, M(J): Next: End Sub
Sub PushSrkey(O() As Srkey, M As Srkey): Dim N&: N = SiSrkey(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiSrkey&(A() As Srkey): On Error Resume Next: SiSrkey = UBound(A) + 1: End Function
Function UbSrkey&(A() As Srkey): UbSrkey = SiSrkey(A) - 1: End Function

Function CiySrkey(K() As Srkey) As Integer()
Dim J%: For J = 0 To UbSrkey(K)
    PushI CiySrkey, K(J).Ci
Next
End Function

Function IsEqSrkey(A As Srkey, B As Srkey) As Boolean
With A
    If .Ci <> B.Ci Then Exit Function
    If .IsDes <> B.IsDes Then Exit Function
End With
IsEqSrkey = True
End Function
