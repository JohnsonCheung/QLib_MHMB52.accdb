Attribute VB_Name = "MxXls_Fea_Lof_Lo_Lof"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Fea_Lof_Lo_Lof."
Public Const LofBdrss$ = ""
'--
Private Type LyUdt
    AliLy() As String
End Type
'--
Enum eLofali: eLofaliL:  eLofaliC:   eLofaliR:   End Enum
Enum eLofBdr: eLofBdrL:  eLofBdrC:   eLofBdrR:   End Enum
Enum eLofagr: eLofagrAvg: eLofagrCnt: eLofagrSum: End Enum
Enum eLofcor: eLofcorOrange: eLofcorGreen: eLofcorLightBlue: End Enum
Public Const SsEnmLofBdr$ = "LofBdrL LofBdrC LofBdrR"
Public Const SsEnmLofagr$ = "LofagrAvg LofagrCnt LofagrSum"
Public Const SsEnmLofali$ = "LofaliL LofaliC LofaliR"
Type Lofali: Fny() As String: Ali As eLofali: End Type 'Deriving(Ay Ctor)
Type Lofwdt: Fny() As String: Wdt As Integer: End Type 'Deriving(Ay Ctor)
Type Lofbdr: Fny() As String: Bdr As eLofBdr: End Type 'Deriving(Ay Ctor)
Type Loflvl: Fny() As String: Lvl As Byte:    End Type 'Deriving(Ay Ctor)
Type Lofcor: Fny() As String: Cor As Long:    End Type 'Deriving(Ay Ctor)
Type Lofagr: Fny() As String: Agr As eLofagr: End Type 'Deriving(Ay Ctor)
Type Loffmt: Fny() As String: Fmt As String:  End Type 'Deriving(Ay Ctor)
Type Loftit: Fldn   As String: Tit() As String:  End Type 'Deriving(Ay Ctor)
Type Loffml: Fldn   As String: Fml As String:  End Type 'Deriving(Ay Ctor)
Type Loflbl: Fldn   As String: Lbl As String:  End Type 'Deriving(Ay Ctor)
Type Lofsum: SumFld As String: FmFld As String: ToFld As String: End Type 'Deriving(Ay Ctor)
Type Lofdta
    Lon As String
    Fny() As String
    Ali() As Lofali
    Bdr() As Lofbdr
    Cor() As Lofcor
    Fml() As Loffml
    Fmt() As Loffmt
    Lbl() As Loflbl
    Lvl() As Loflvl
    Sum() As Lofsum
    Tit() As Loftit
    Agr() As Lofagr
    Wdt() As Lofwdt
End Type
Function StrEnmLofali$(E As eLofali)

End Function
Function StrEnmLofBdr$(E As eLofBdr)

End Function
Function StrEnmLofcor$(E As eLofcor)

End Function

Function StrEnmLofagr$(E As eLofagr)

End Function
Function Lofsum(SumFld, FmFld, ToFld) As Lofsum
With Lofsum
    .SumFld = SumFld
    .FmFld = FmFld
    .ToFld = ToFld
End With
End Function
Function LofsumAdd(A As Lofsum, B As Lofsum) As Lofsum(): PushLofSum LofsumAdd, A: PushLofSum LofsumAdd, B: End Function
Sub PushLofSumAy(O() As Lofsum, A() As Lofsum): Dim J&: For J = 0 To LofsumUB(A): PushLofSum O, A(J): Next: End Sub
Sub PushLofSum(O() As Lofsum, M As Lofsum): Dim N&: N = LofsumSI(O): ReDim Preserve O(N): O(N) = M: End Sub
Function LofsumSI&(A() As Lofsum): On Error Resume Next: LofsumSI = UBound(A) + 1: End Function
Function LofsumUB&(A() As Lofsum): LofsumUB = LofsumSI(A) - 1: End Function
Function Loflbl(Fldn, Lbl) As Loflbl
With Loflbl
    .Fldn = Fldn
    .Lbl = Lbl
End With
End Function
Function LoflblAdd(A As Loflbl, B As Loflbl) As Loflbl(): PushLoflbl LoflblAdd, A: PushLoflbl LoflblAdd, B: End Function
Sub PushLoflbly(O() As Loflbl, A() As Loflbl): Dim J&: For J = 0 To LoflblUB(A): PushLoflbl O, A(J): Next: End Sub
Sub PushLoflbl(O() As Loflbl, M As Loflbl): Dim N&: N = LoflblSI(O): ReDim Preserve O(N): O(N) = M: End Sub
Function LoflblSI&(A() As Loflbl): On Error Resume Next: LoflblSI = UBound(A) + 1: End Function
Function LoflblUB&(A() As Loflbl): LoflblUB = LoflblSI(A) - 1: End Function
Function Loffml(Fldn, Fml) As Loffml
With Loffml
    .Fldn = Fldn
    .Fml = Fml
End With
End Function
Function LoffmlAdd(A As Loffml, B As Loffml) As Loffml(): PushLoffml LoffmlAdd, A: PushLoffml LoffmlAdd, B: End Function
Sub PushLoffmlAy(O() As Loffml, A() As Loffml): Dim J&: For J = 0 To LoffmlUB(A): PushLoffml O, A(J): Next: End Sub
Sub PushLoffml(O() As Loffml, M As Loffml): Dim N&: N = LoffmlSI(O): ReDim Preserve O(N): O(N) = M: End Sub
Function LoffmlSI&(A() As Loffml): On Error Resume Next: LoffmlSI = UBound(A) + 1: End Function
Function LoffmlUB&(A() As Loffml): LoffmlUB = LoffmlSI(A) - 1: End Function
Function Loftit(Fldn, Tit) As Loftit
With Loftit
    .Fldn = Fldn
    .Tit = Tit
End With
End Function
Function AddLoftit(A As Loftit, B As Loftit) As Loftit(): PushLoftit AddLoftit, A: PushLoftit AddLoftit, B: End Function
Sub PushLoftity(O() As Loftit, A() As Loftit): Dim J&: For J = 0 To LoftitUB(A): PushLoftit O, A(J): Next: End Sub
Sub PushLoftit(O() As Loftit, M As Loftit): Dim N&: N = LoftitSI(O): ReDim Preserve O(N): O(N) = M: End Sub
Function LoftitSI&(A() As Loftit): On Error Resume Next: LoftitSI = UBound(A) + 1: End Function
Function LoftitUB&(A() As Loftit): LoftitUB = LoftitSI(A) - 1: End Function
Function Loffmt(Fny$(), Fmt) As Loffmt
With Loffmt
    .Fny = Fny
    .Fmt = Fmt
End With
End Function
Function LoffmtAdd(A As Loffmt, B As Loffmt) As Loffmt(): PushLoffmt LoffmtAdd, A: PushLoffmt LoffmtAdd, B: End Function
Sub PushLoffmtAy(O() As Loffmt, A() As Loffmt): Dim J&: For J = 0 To UbLoffmt(A): PushLoffmt O, A(J): Next: End Sub
Sub PushLoffmt(O() As Loffmt, M As Loffmt): Dim N&: N = SiLoffmt(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiLoffmt&(A() As Loffmt): On Error Resume Next: SiLoffmt = UBound(A) + 1: End Function
Function UbLoffmt&(A() As Loffmt): UbLoffmt = SiLoffmt(A) - 1: End Function
Function Lofagr(Fny$(), Agr As eLofagr) As Lofagr
With Lofagr
    .Fny = Fny
    .Agr = Agr
End With
End Function
Function LofagrAdd(A As Lofagr, B As Lofagr) As Lofagr(): PushLofagr LofagrAdd, A: PushLofagr LofagrAdd, B: End Function
Sub PushLofagry(O() As Lofagr, A() As Lofagr): Dim J&: For J = 0 To LofagrUB(A): PushLofagr O, A(J): Next: End Sub
Sub PushLofagr(O() As Lofagr, M As Lofagr): Dim N&: N = LofagrSI(O): ReDim Preserve O(N): O(N) = M: End Sub
Function LofagrSI&(A() As Lofagr): On Error Resume Next: LofagrSI = UBound(A) + 1: End Function
Function LofagrUB&(A() As Lofagr): LofagrUB = LofagrSI(A) - 1: End Function
Function Lofcor(Fny$(), Cor) As Lofcor
With Lofcor
    .Fny = Fny
    .Cor = Cor
End With
End Function
Function LofcorAdd(A As Lofcor, B As Lofcor) As Lofcor(): PushLofcor LofcorAdd, A: PushLofcor LofcorAdd, B: End Function
Sub PushLofcorAy(O() As Lofcor, A() As Lofcor): Dim J&: For J = 0 To LofcorUB(A): PushLofcor O, A(J): Next: End Sub
Sub PushLofcor(O() As Lofcor, M As Lofcor): Dim N&: N = LofcorSI(O): ReDim Preserve O(N): O(N) = M: End Sub
Function LofcorSI&(A() As Lofcor): On Error Resume Next: LofcorSI = UBound(A) + 1: End Function
Function LofcorUB&(A() As Lofcor): LofcorUB = LofcorSI(A) - 1: End Function
Function Loflvl(Fny$(), Lvl) As Loflvl
With Loflvl
    .Fny = Fny
    .Lvl = Lvl
End With
End Function
Function LoflvlAdd(A As Loflvl, B As Loflvl) As Loflvl(): PushLoflvl LoflvlAdd, A: PushLoflvl LoflvlAdd, B: End Function
Sub PushLoflvlAy(O() As Loflvl, A() As Loflvl): Dim J&: For J = 0 To LoflvlUB(A): PushLoflvl O, A(J): Next: End Sub
Sub PushLoflvl(O() As Loflvl, M As Loflvl): Dim N&: N = LoflvlSI(O): ReDim Preserve O(N): O(N) = M: End Sub
Function LoflvlSI&(A() As Loflvl): On Error Resume Next: LoflvlSI = UBound(A) + 1: End Function
Function LoflvlUB&(A() As Loflvl): LoflvlUB = LoflvlSI(A) - 1: End Function
Function Lofbdr(Fny$(), Bdr As eLofBdr) As Lofbdr
With Lofbdr
    .Fny = Fny
    .Bdr = Bdr
End With
End Function
Function LofbdrAdd(A As Lofbdr, B As Lofbdr) As Lofbdr(): PushLofBdr LofbdrAdd, A: PushLofBdr LofbdrAdd, B: End Function
Sub PushLofBdrAy(O() As Lofbdr, A() As Lofbdr): Dim J&: For J = 0 To LofBdrUB(A): PushLofBdr O, A(J): Next: End Sub
Sub PushLofBdr(O() As Lofbdr, M As Lofbdr): Dim N&: N = LofBdrSI(O): ReDim Preserve O(N): O(N) = M: End Sub
Function LofBdrSI&(A() As Lofbdr): On Error Resume Next: LofBdrSI = UBound(A) + 1: End Function
Function LofBdrUB&(A() As Lofbdr): LofBdrUB = LofBdrSI(A) - 1: End Function
Function Lofwdt(Fny$(), Wdt) As Lofwdt
With Lofwdt
    .Fny = Fny
    .Wdt = Wdt
End With
End Function
Function LofwdtAdd(A As Lofwdt, B As Lofwdt) As Lofwdt(): PushLofwdt LofwdtAdd, A: PushLofwdt LofwdtAdd, B: End Function
Sub PushLofwdty(O() As Lofwdt, A() As Lofwdt): Dim J&: For J = 0 To LofwdtUB(A): PushLofwdt O, A(J): Next: End Sub
Sub PushLofwdt(O() As Lofwdt, M As Lofwdt): Dim N&: N = LofwdtSI(O): ReDim Preserve O(N): O(N) = M: End Sub
Function LofwdtSI&(A() As Lofwdt): On Error Resume Next: LofwdtSI = UBound(A) + 1: End Function
Function LofwdtUB&(A() As Lofwdt): LofwdtUB = LofwdtSI(A) - 1: End Function
Function Lofali(Fny$(), Ali As eLofali) As Lofali
With Lofali
    .Fny = Fny
    .Ali = Ali
End With
End Function
Function LofaliAdd(A As Lofali, B As Lofali) As Lofali(): PushLofali LofaliAdd, A: PushLofali LofaliAdd, B: End Function
Sub PushLofaliAy(O() As Lofali, A() As Lofali): Dim J&: For J = 0 To LofaliUB(A): PushLofali O, A(J): Next: End Sub
Sub PushLofali(O() As Lofali, M As Lofali): Dim N&: N = LofaliSI(O): ReDim Preserve O(N): O(N) = M: End Sub
Function LofaliSI&(A() As Lofali): On Error Resume Next: LofaliSI = UBound(A) + 1: End Function
Function LofaliUB&(A() As Lofali): LofaliUB = LofaliSI(A) - 1: End Function

Private Function VVLofu(Lof$()) As Lofdta
Dim L As LyUdt: L = W2LyUdt(Lof)
Dim S As TSpec
With VVLofu
'    .Fny = A.Fny
'    .Lon = A.Lon
    .Ali = W2Lofali(LyTSpeci(S, "Ali"))
'    .Wdt = W2Lofwdt(LyTSpeci(S, "Wdt"))
'    .Bdr = W2LofBdr(LyTSpeci(S, "bdr"))
'    .Lvl = W2Loflvl(LyTSpeci(S, "Lvl"))
'    .Cor = W2Lofcor(LyTSpeci(S, "Cor"))
'    .Tot = W2Loftot(LyTSpeci(S, "Tot"))
'    .Fmt = W2Loffmt(LyTSpeci(S, "Fmt"))
'    .Tit = W2LofTit(LyTSpeci(S, "Tit"))
'    .Fml = W2Loffml(LyTSpeci(S, "Fml"))
'    .Lbl = W2Loflbl(LyTSpeci(S, "Lbl"))
'    .Sum = W2LofSum(LyTSpeci(S, "Sum"))
End With
End Function
Private Function W2LyUdt(Lofly$()) As LyUdt

End Function
Private Function W2Lofali(Ali$()) As Lofali()
Dim L: For Each L In Itr(Ali)
    PushLofali W2Lofali, W2LofOneAli(L)
Next
End Function
Private Function W2LofOneAli(Ali) As Lofali
End Function

Function LofTitAy(LofTitLy$()) As Loftit()
Dim L: For Each L In Itr(LofTitLy)
    PushLoftit LofTitAy, LofTitLn(L)
Next
End Function

Function LofTitLn(LofTitLin) As Loftit
Dim O As Loftit
With BrkSpc(LofTitLin)
    O.Fldn = .S1
    O.Tit = AmTrim(SplitVBar(.S2))
End With
End Function

Function LofSampFny() As String()
LofSampFny = SySs("A B C D E F")
End Function

Sub LofSampBrw(): Brw LofUdFmt(LofUdSamp): End Sub

Function LofSamp() As String()
ClrBfr
BfrV "Ali Center F"
BfrV "Ali Left B"
BfrV "Ali Right D E"
BfrV "Bdr Center F"
BfrV "Bdr Left"
BfrV "Bdr Right G"
BfrV "Cor 12345 B"
BfrV "Fml C B * 2"
BfrV "Fml F A + B"
BfrV "Fmt #,## B C"
BfrV "Fmt #,##.## D E"
BfrV "Lbl A lksd flks dfj"
BfrV "Lbl B lsdkf lksdf klsdj f"
BfrV "Lbl A lksd flks dfj"
BfrV "Lo Fld B C D E F G"
BfrV "Lo Nm BC"
BfrV "Lvl 2 C"
BfrV "Sum A B X"
BfrV "Tit A bc | sdf"
BfrV "Tit B bc | sdkf | sdfdf"
BfrV "Tot Avg D"
BfrV "Tot Cnt C"
BfrV "Tot Sum B"
BfrV "Wdt 10 B X"
BfrV "Wdt 20 D C C"
BfrV "Wdt 3000 E F G C"
LofSamp = FmtT2ry(LyBfr)
End Function
