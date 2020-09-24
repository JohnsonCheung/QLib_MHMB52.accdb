Attribute VB_Name = "MxXls_Fea_Lof_Fmt"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Fea_Lof_Fmt."
Const LoflnTpLon$ = "Lon ?"
Const LoflnTpFny$ = "Fny ?"
Const LoflnTpAli$ = "Ali ? ?"
Const LoflnTpBdr$ = "Bdr ? ?"
Const LoLnTpCor$ = "Cor ? ?"
Const LoflnTpFml$ = "Fml ? ?"
Const LoflnTpFmt$ = "Fmt ? ?"
Const LoflnTpLbl$ = "Lbl ? ?"
Const LoflnTpLvl$ = "Lvl ? ?"
Const LoflnTpSum$ = "Sum ? ?"
Const LoflnTpTit$ = "Tit ? ?"
Const LoflnTpAgr$ = "Tot ? ?"
Const LoflnTpWdt$ = "Wdt ? ?"

Private Sub B_LofUdFmt(): Brw LofUdFmt(LofUdSamp): End Sub
Function LofUdFmt(A As Lofdta) As String()
Dim O$()
With A
PushI O, W_Lon(.Lon)
PushI O, W_Fny(.Fny)
PushIAy O, W_Ali(.Ali)
PushIAy O, W_Bdr(.Bdr)
PushIAy O, W_Cor(.Cor)
PushIAy O, W_Fml(.Fml)
PushIAy O, W_Fmt(.Fmt)
PushIAy O, W_Lbl(.Lbl)
PushIAy O, W_Lvl(.Lvl)
PushIAy O, W_Sum(.Sum)
PushIAy O, W_Tit(.Tit)
PushIAy O, W_Agr(.Agr)
PushIAy O, W_Wdt(.Wdt)
End With
LofUdFmt = O
End Function

Private Function W_Lon$(Lon$):   W_Lon = FmtQQ(LoflnTpLon, Lon):        End Function
Private Function W_Fny$(Fny$()): W_Fny = FmtQQ(LoflnTpFny, JnSpc(Fny)): End Function

Private Function W_Ali(A() As Lofali) As String(): Dim J%: For J = 0 To LofaliUB(A): PushI W_Ali, W_ItmAli(A(J)): Next: End Function
Private Function W_Bdr(A() As Lofbdr) As String(): Dim J%: For J = 0 To LofBdrUB(A): PushI W_Bdr, W_ItmBdr(A(J)): Next: End Function
Private Function W_Cor(A() As Lofcor) As String(): Dim J%: For J = 0 To LofcorUB(A): PushI W_Cor, W_ItmCor(A(J)): Next: End Function
Private Function W_Fml(A() As Loffml) As String(): Dim J%: For J = 0 To LoffmlUB(A): PushI W_Fml, W_ItmFml(A(J)): Next: End Function
Private Function W_Fmt(A() As Loffmt) As String(): Dim J%: For J = 0 To UbLoffmt(A): PushI W_Fmt, W_ItmFmt(A(J)): Next: End Function
Private Function W_Lbl(A() As Loflbl) As String(): Dim J%: For J = 0 To LoflblUB(A): PushI W_Lbl, W_ItmLbl(A(J)): Next: End Function
Private Function W_Lvl(A() As Loflvl) As String(): Dim J%: For J = 0 To LoflvlUB(A): PushI W_Lvl, W_ItmLvl(A(J)): Next: End Function
Private Function W_Sum(A() As Lofsum) As String(): Dim J%: For J = 0 To LofsumUB(A): PushI W_Sum, W_ItmSum(A(J)): Next: End Function
Private Function W_Tit(A() As Loftit) As String(): Dim J%: For J = 0 To LoftitUB(A): PushI W_Tit, W_ItmTit(A(J)): Next: End Function
Private Function W_Agr(A() As Lofagr) As String(): Dim J%: For J = 0 To LofagrUB(A): PushI W_Agr, W_ItmTot(A(J)): Next: End Function
Private Function W_Wdt(A() As Lofwdt) As String(): Dim J%: For J = 0 To LofwdtUB(A): PushI W_Wdt, W_ItmWdt(A(J)): Next: End Function

Private Function W_ItmAli$(A As Lofali): W_ItmAli = FmtQQ(LoflnTpBdr, StrEnmLofali(A.Ali), JnSpc(A.Fny)): End Function
Private Function W_ItmBdr$(A As Lofbdr): W_ItmBdr = FmtQQ(LoflnTpBdr, StrEnmLofBdr(A.Bdr), JnSpc(A.Fny)): End Function
Private Function W_ItmCor$(A As Lofcor): W_ItmCor = FmtQQ(LoLnTpCor, StrEnmLofcor(A.Cor), JnSpc(A.Fny)):  End Function
Private Function W_ItmFml$(A As Loffml): W_ItmFml = FmtQQ(LoflnTpFml, A.Fldn, A.Fml):                     End Function
Private Function W_ItmFmt$(A As Loffmt): W_ItmFmt = FmtQQ(LoflnTpFmt, A.Fmt, JnSpc(A.Fny)):               End Function
Private Function W_ItmLbl$(A As Loflbl): W_ItmLbl = FmtQQ(LoflnTpLbl, A.Fldn, A.Lbl):                     End Function
Private Function W_ItmLvl$(A As Loflvl): W_ItmLvl = FmtQQ(LoflnTpLvl, A.Lvl, JnSpc(A.Fny)):               End Function
Private Function W_ItmSum$(A As Lofsum): W_ItmSum = FmtQQ(LoflnTpSum, A.SumFld, A.FmFld, A.ToFld):        End Function
Private Function W_ItmTit$(A As Loftit): W_ItmTit = FmtQQ(LoflnTpBdr, A.Fldn, A.Tit):                     End Function
Private Function W_ItmTot$(A As Lofagr): W_ItmTot = FmtQQ(LoflnTpBdr, StrEnmLofagr(A.Agr), JnSpc(A.Fny)): End Function
Private Function W_ItmWdt$(A As Lofwdt): W_ItmWdt = FmtQQ(LoflnTpBdr, A.Wdt, JnSpc(A.Fny)):               End Function

Function LofUdSamp() As Lofdta: LofUdSamp = LofUd1(LofSamp): End Function
