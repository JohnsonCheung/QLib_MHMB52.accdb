Attribute VB_Name = "MxDta_Da_Brw"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Brw."

Sub BrwDrs2(A As Drs, B As Drs, Optional NN$ = "Drs1 Drs2", Optional Tit$ = "Brw 2 Drs")
Dim AyA$(), AyB$(), N1$, N2$
N1 = BefSpc(NN)
N2 = AftSpc(NN)
AyA = FmtDrs(A)
AyB = FmtDrs(B)
BrwAy SyAddAp(Box(Tit), AyA, AyB), "BrwDrs2_"
End Sub

Sub BrwDrs3(A As Drs, B As Drs, C As Drs, Optional ByVal NN$ = "Drs1 Drs2 Drs3", Optional Tit$ = "Brw 3 Drs")
Dim AyA$(), AyB$(), AyC$(), N1$, N2$, N3$
N1 = ShfTm(NN)
N2 = ShfTm(NN)
N3 = NN
AyA = SyAdd(Box(N1), FmtDrs(A))
AyB = SyAdd(Box(N2), FmtDrs(B))
AyC = SyAdd(Box(N3), FmtDrs(C))
BrwAy SyAddAp(Box(Tit), AyA, AyB, AyC), "BrwDrs3_"
End Sub

Sub BrwDrs4(A As Drs, B As Drs, C As Drs, D As Drs, Optional ByVal NN$ = "Drs1 Drs2 Drs3 Drs4", Optional Tit$ = "Brw 4 Drs")
Dim AyA$(), AyB$(), AyC$(), AyD$(), N1$, N2$, N3$, N4$
N1 = ShfTm(NN)
N2 = ShfTm(NN)
N3 = ShfTm(NN)
N4 = NN
AyA = SyAdd(Box(N1), FmtDrs(A))
AyB = SyAdd(Box(N2), FmtDrs(B))
AyC = SyAdd(Box(N3), FmtDrs(C))
AyD = SyAdd(Box(N4), FmtDrs(D))
BrwAy SyAddAp(Box(Tit), AyA, AyB, AyC, AyD), "BrwDrs4_"
End Sub

Sub DmpDrs(D As Drs, _
Optional Rds As eRds, _
Optional Fmt As eTblFmt, _
Optional Zer As eZer, _
Optional Wdt% = 100, _
Optional ColnnAliR$, _
Optional ColnnSum$, _
Optional NoIx As Boolean, _
Optional PfxFn$ = "Drs_")
Dim F$(): F = FmtDrs(D, Rds, Fmt, Zer, Wdt, ColnnAliR, ColnnSum, NoIx)
DmpAy F
End Sub
Sub BrwDrs(D As Drs, _
Optional Rds As eRds, _
Optional Fmt As eTblFmt, _
Optional Zer As eZer, _
Optional Wdt%, _
Optional ColnnAliR$, _
Optional ColnnSum$, _
Optional NoIx As Boolean, _
Optional PfxFn$ = "Drs_")
BrwAy FmtDrs(D, Rds, Fmt, Zer, Wdt, ColnnAliR, ColnnSum, NoIx), PfxFn
End Sub
Sub VcDrs(D As Drs, _
Optional Rds As eRds, _
Optional Fmt As eTblFmt, _
Optional Zer As eZer, _
Optional Wdt% = 100, _
Optional ColnnAliR$, _
Optional ColnnSum$, _
Optional NoIx As Boolean, _
Optional PfxFn$ = "Drs_")
VcAy FmtDrs(D, Rds, Fmt, Zer, Wdt, ColnnAliR, ColnnSum, NoIx), PfxFn
End Sub

Sub BrwDy(D(), Optional Fmt As eTblFmt, Optional Zer As eZer, Optional Wdt%, Optional AlirCii$, Optional NoIx As Boolean)
BrwAy FmtDy(D, Fmt, Zer, Wdt, AlirCii, NoIx)
End Sub

Sub DmpDy(D(), Optional Fmt As eTblFmt, Optional Zer As eZer, Optional Wdt%, Optional AlirCii$, Optional NoIx As Boolean)
BrwAy FmtDy(D, Fmt, Zer, Wdt, AlirCii, NoIx)
End Sub
