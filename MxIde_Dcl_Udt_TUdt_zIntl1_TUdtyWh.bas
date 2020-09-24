Attribute VB_Name = "MxIde_Dcl_Udt_TUdt_zIntl1_TUdtyWh"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_Udt_TUdt_TUdtOp."
Private Type TWhUdt: WhPPrv As eWhPPrv: RxAyMdn() As RegExp: RxAyUdtn() As RegExp: End Type
Private Sub B_TUdtyWh()
GoSub ZZ
Exit Sub
Dim LpmWhTUdt$, Act() As TUdt
ZZ:
    LpmWhTUdt = "-Mdn Db"
    Act = TUdtyWh(TUdtyPC, LpmWhTUdt)
    Brw FmtT3ry(UdtilnyTUdty(Act))
    Return
End Sub
Function TUdtyWh(U() As TUdt, LpmWhTUdt$) As TUdt()
If LpmWhTUdt = "" Then
    TUdtyWh = U
    Exit Function
End If
Dim Wh As TWhUdt: Wh = TWhUdt(LpmWhTUdt)
Dim J%: For J = 0 To UbTUdt(U)
    If IsWhTUdt(U(J), Wh) Then
        'Debug.Print U(J).Mdn & "." & U(J).Udtn
        PushTUdt TUdtyWh, U(J) '<===
    End If
Next
End Function
Private Function IsWhTUdt(U As TUdt, Wh As TWhUdt) As Boolean
With Wh
    If Not HasRxAyAnd(U.Mdn, .RxAyMdn) Then Exit Function
    If Not HasRxAyAnd(U.Udtn, .RxAyUdtn) Then Exit Function
    If Not HitPPrv(U.IsPrv, .WhPPrv) Then Exit Function
End With
IsWhTUdt = True
End Function
Private Function TWhUdt(LpmWhTUdt$) As TWhUdt
Dim P As TLpmBrk: P = TLpmBrk(LpmWhTUdt)
Dim O As TWhUdt
    Dim IsPub As Boolean, IsPrv As Boolean
    Dim RxAyMdn() As RegExp, RxAyUdtn() As RegExp
    IsPub = SwvLpm(P, "Pub")
    IsPrv = SwvLpm(P, "Prv")
    RxAyMdn = RxayLpm(P, "Mdn")
    RxAyUdtn = RxayLpm(P, "Udtn")

With TWhUdt
    .RxAyMdn = RxAyMdn
    .RxAyUdtn = RxAyUdtn
    .WhPPrv = eWhPPrvIsPPrv2(IsPub, IsPrv)
End With
End Function
