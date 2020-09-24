Attribute VB_Name = "MxDao_Sql_Fmt_zIntl_TQpFun"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Sql_TQp_TQpOp."

Function TQpQp(Qp) As TQp
Dim IQp$: IQp = Qp
With TQpQp
    .Qpt = ShfQpt(IQp)
    .Qpr = IQp
End With
End Function
Private Function ShfQpt(OQpKwwAdjd$) As eQpt
Dim QpSav$: QpSav = OQpKwwAdjd
'Public Const TmlKwwQp$ = "From [Group By] Having [Inner Join] [Left Join] [Order By] Select [Select Distinct] Set Update Where"
Dim Kwwy$(): Kwwy = Qpkwwy
Dim O As eQpt
Dim Kww$: Kww = ShfPfxySpc(OQpKwwAdjd, Kwwy, eCasIgn)
Select Case Kww
Case "From":     O = eQptFm
Case "Group By": O = eQptGp
Case "Having":   O = eQptHav
Case "Inner Join": O = eQptInrJn
Case "Left Join": O = eQptLeftJn
Case "Update":   O = eQptUpd
Case "Set":  O = eQptSet
Case "Select Distinct": O = eQptSelDis
Case "Select": O = eQptSel
Case "Where": O = eQptWh
Case "Order By": O = eQptOrd
Case "Into": O = eQptInto
Case Else
    Stop
    ThwPm CSub, "OQpKwwAdjd has invalid Kww", "OQpKwwAdjd", QpSav
End Select
ShfQpt = O
End Function
Private Sub B_TQpyQpy()
GoSub T1
Exit Sub
Dim Qpy$()
T1:
    Erase XX
    X "From ((((((OH       As x"
    X "Inner Join YMDCur   As z On (x.DD         = z.DD) AND (x.MM=z.MM) AND (x.YY=z.YY))"
    X "Left Join YpStk     As a On x.YpStk       = a.YpStk)"
    X "Left Join q1SKU     As b On x.Sku         = b.Sku)"
    X "Left Join SHBrandQ  As c On b.CdB         = c.CdQly)"
    X "Left Join SHBrand   As d On c.CdSHBrand   = d.CdSHBrand)"
    X "Left Join FinStream As e On d.CdFinStream = e.CdFinStream)"
    X "Left Join Hse       As f On c.Hse         = f.Hse"
    X "Order By SnoFinStream, NmFinStream, SnoHse, NmHse, SnoSHBrand, NmSHBrand, b.NmB;"
    Qpy = XX
Tst:
    Dim Q() As TQp: Q = TQpyQpy(Qpy)
    Dim Qp$(): Qp = QpyTQpy(Q)
    If Not IsEqAy(Qp, Qpy) Then Stop
    Return
End Sub
Function QpyTQpy(Q() As TQp) As String()
Dim J%: For J = 0 To UbTQp(Q)
    PushI QpyTQpy, EnmtQpt(Q(J).Qpt) & " " & Q(J).Qpr
Next
End Function
Function TQpyQpy(Qpy$()) As TQp()
Dim Qp: For Each Qp In Itr(Qpy)
    PushTQp TQpyQpy, TQpQp(Qp)
Next
End Function
Function eQptyTQpy(Q() As TQp) As eQpt()
Dim J%: For J = 0 To UbTQp(Q)
    PushI eQptyTQpy, Q(J).Qpt
Next
End Function
Function eQptQp(Qp) As eQpt: eQptQp = EnmvQpt(PfxPfxySpc(Qp, Qpkwwy)): End Function
Function TQpyRplBix(Q() As TQp, Bix&, By() As TQp) As TQp()
Dim O() As TQp: O = Q
Dim J%: For J = Bix To Bix + UbTQp(By)
    O(J) = By(J - Bix)
Next
TQpyRplBix = O
End Function
Function TQpyRplIy(Q() As TQp, Iy%(), By() As TQp) As TQp()
If Si(Iy) <> SiTQp(By) Then ThwPm CSub, "@TQpBy & @Iy does not have same si", "TQpBy-Si Biy-Si", SiTQp(Q), Si(Iy)
Dim O() As TQp: O = Q
Dim J%: For J = 0 To UB(Iy)
    Dim I%: I = Iy(J)
    O(I) = By(J)
Next
TQpyRplIy = O
End Function

Function TQpyWhBei(Q() As TQp, B As Bei) As TQp()
Dim J%: For J = B.Bix To B.Eix
    PushTQp TQpyWhBei, Q(J)
Next
End Function
Function TQpyWhIy(Q() As TQp, Iy%()) As TQp()
Dim J%: For J = 0 To UB(Iy)
    PushTQp TQpyWhIy, Q(Iy(J))
Next
End Function
Function SqlTQpy$(Q() As TQp)
Dim O$()
Dim J%: For J = 0 To UbTQp(Q)
    PushI O, StrTQp(Q(J))
Next
SqlTQpy = JnCrLf(LyEndTrim(O))
End Function

