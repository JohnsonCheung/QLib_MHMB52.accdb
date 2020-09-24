Attribute VB_Name = "MxDao_Sql_Fmt_zIntl1_P12Kww"
Option Compare Database
Option Explicit

Private Sub B_P12Kwy()
GoSub T1
Exit Sub
Dim S, Kwy$(), C As eCas, OPosBeg%, OPosBegEpt%, Act As P12, Ept As P12
T1:
        '123456789012345
    S = "123 ABC    DEF xx ABC   DEF xx"
    Kwy = SySs("ABC DEF")
    OPosBegEpt = 15
    Ept = P12(5, 14)
    OPosBeg = 1
    GoTo Tst
Tst:
    Act = P12Kwy(S, Kwy, C, OPosBeg)
    Debug.Assert IsEqP12(Act, Ept)
    Debug.Assert OPosBegEpt = OPosBeg
    Return
End Sub

Private Sub B_P12yKwy()
'GoSub T1
GoSub T2
Exit Sub
Dim S, Kwy$(), C As eCas, Act() As P12, Ept() As P12
T1:     '         1         2
        '123456789012345678901234567890
    S = "123 ABC    DEF xx ABC   DEF xx"
    Kwy = SySs("ABC DEF")
    PushP12 Ept, P12(5, 14)
    PushP12 Ept, P12(19, 27)
    GoTo Tst
T2:
    S = "SELECT x.YY, x.MM, x.DD, NmYpStk, Bott, [Bott]*[Bott2SC] AS StdCase, Size, StdCaseSize, x.Val, x.Sku, DesSKU, b.CdB AS [Quality Code], b.NmB AS Quality, d.CdSHBrand AS [Brand Code], NmSHBrand AS Brand, NmFinStream AS Stream, NmHse AS House, SnoHse, SnoFinStream, SnoSHBrand  FROM ((((((OH AS x  Inner Join YMDCur AS z ON (x.DD=z.DD) AND (x.MM=z.MM) AND (x.YY=z.YY)) LEFT JOIN YpStk AS a ON x.YpStk       = a.YpStk) LEFT JOIN q1SKU AS b ON x.Sku         = b.Sku) LEFT JOIN SHBrandQ AS c ON b.CdB         = c.CdQly) LEFT JOIN SHBrand AS d ON c.CdSHBrand   = d.CdSHBrand) LEFT JOIN FinStream AS e ON d.CdFinStream = e.CdFinStream) LEFT JOIN Hse AS f ON c.Hse         = f.Hse  ORDER BY SnoFinStream, NmFinStream, SnoHse, NmHse, SnoSHBrand, NmSHBrand, b.NmB;  "
    Kwy = SySs("Left Join")
    GoTo Tst
Tst:
    Act = P12yKwy(S, Kwy, C)
    Debug.Assert IsEqP12y(Act, Ept)
    Return
End Sub
Private Function P12Kwy(Sql, Kwy, C As eCas, OPosBeg%) As P12
Dim P%, PosBeg%
Dim P1%, P2%
Dim U%: U = UB(Kwy)
PosBeg = OPosBeg
Dim Kw$, I%: For I = 0 To UB(Kwy)
    Kw = Kwy(I)
    P = PosSsub(Sql, Kw + " ", C, PosBeg): If P = 0 Then Exit Function
    If I = 0 Then P1 = P
    If I = U Then P2 = P + Len(Kwy(I)) - 1: Exit For
    PosBeg = PosNxtKw(Sql, P, Kwy(I))
Next
P12Kwy = P12(P1, P2)
OPosBeg = P2 + 1
End Function
Private Function PosNxtKw%(S, PosKw%, Kw)
Dim AftKw$: AftKw = Mid(S, PosKw)
Dim NSpc%: NSpc = NSpcPfx(AftKw)
PosNxtKw = PosKw + Len(Kw) + NSpc
End Function
Function P12yKwy(S, Kwy, Optional C As eCas) As P12()
Dim PosBeg%: PosBeg = 1
Dim Cnt&, L%: L = Len(S)
Again:
    RaiseLoopTooMuch CSub, Cnt, 20
    If PosBeg >= L Then Exit Function
    Dim P As P12: P = P12Kwy(S, Kwy, C, PosBeg)
    If P.P1 = 0 Then Exit Function
    PushP12 P12yKwy, P
    GoTo Again
End Function


