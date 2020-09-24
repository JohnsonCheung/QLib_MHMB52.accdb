Attribute VB_Name = "MxIde_Dcl_Udt_TUdt"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_Udt_TUdt."
Type TUmb: Mbn As String: IsAy As Boolean: Tyn As String: Rmky() As String: End Type 'Deriving(Ctor Ay)
Type TUdt
    Mdn As String
    IsPrv As Boolean
    Udtn As String
    Mbr() As TUmb
    GenCtor As Boolean
    GenAy As Boolean
    GenAdd As Boolean
    GenPushAy As Boolean
    GenOpt As Boolean
    Rmky() As String ' It comes from the rmk of lasLn aft rmv the Deriving(...)
End Type 'Deriving(Ay Ctor)

Function TUmb(IsAy As Boolean, Mbn$, Tyn$) As TUmb
With TUmb
    .IsAy = IsAy
    .Mbn = Mbn
    .Tyn = Tyn
End With
End Function
Sub PushTUmb(O() As TUmb, M As TUmb)
Dim N&: N = SiTUdtMbr(O)
ReDim Preserve O(N)
O(N) = M
End Sub

Function TUdt(Mdn, IsPrv, Udtn, Mbr() As TUmb, GenCtor, GenAy, GenOpt, GenPushAy, GenAdd) As TUdt
With TUdt
    .IsPrv = IsPrv
    .Udtn = Udtn
    .Mbr = Mbr
    .GenCtor = GenCtor
    .GenAy = GenAy
    .GenOpt = GenOpt
    .GenPushAy = GenPushAy
    .GenAdd = GenAdd
End With
End Function
Function SiTUdtMbr&(A() As TUmb): On Error Resume Next: SiTUdtMbr = UBound(A) + 1: End Function
Function UbTUmb&(A() As TUmb): UbTUmb = SiTUdtMbr(A) - 1: End Function
Function SiTUdt&(A() As TUdt): On Error Resume Next: SiTUdt = UBound(A) + 1: End Function
Function UbTUdt&(A() As TUdt): UbTUdt = SiTUdt(A) - 1: End Function
Sub PushTUdt(O() As TUdt, M As TUdt)
Dim N&: N = SiTUdt(O)
ReDim Preserve O(N)
O(N) = M
End Sub
Sub PushTUdty(O() As TUdt, M() As TUdt)
Dim J%: For J = 0 To UbTUdt(M)
    PushTUdt O, M(J)
Next
End Sub
Function TUdtySng(U As TUdt) As TUdt(): PushTUdt TUdtySng, U: End Function
Function SampTUdt() As TUdt
Dim M() As TUmb
    PushTUmb M, TUmb(True, "MbrA", "String")
    PushTUmb M, TUmb(True, "MbrB", "String")
    PushTUmb M, TUmb(True, "MbrC", "Worksheet")
SampTUdt = TUdt("MdnXXX", True, "UdtA", M, True, True, True, True, True)
End Function

