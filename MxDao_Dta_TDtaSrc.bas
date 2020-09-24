Attribute VB_Name = "MxDao_Dta_TDtaSrc"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Dta_TDtaSrc_Udt."
Type TF: Tbn As String: Fny() As String: End Type 'Deriving(Ay Ctor Opt)
Type TDtaSrcFm: TDtaSrcFfn As String: TDtaSrcn As String: End Type 'Deriving(Ctor Opt) TDtaSrcn may be like {Lo | Tbl | ..}
Type TDtaSrcFmOpt: Som As Boolean: TDtaSrcFm As TDtaSrcFm: End Type
Type TDtaSrc: Fm As TDtaSrcFm: TF() As TF: End Type 'Deriving(Ctor)
Type TDtaSrcOpt: Som As Boolean: TDtaSrc As TDtaSrc: End Type
Function TDtaSrcFmOpt(Som, A As TDtaSrcFm) As TDtaSrcFmOpt: With TDtaSrcFmOpt: .Som = Som: .TDtaSrcFm = A: End With: End Function
Function SomTDtaSrcFm(A As TDtaSrcFm) As TDtaSrcFmOpt: SomTDtaSrcFm.Som = True: SomTDtaSrcFm.TDtaSrcFm = A: End Function
Function TDtaSrcFm(TDtaSrcFfn, TDtaSrcn) As TDtaSrcFm
With TDtaSrcFm
    .TDtaSrcFfn = TDtaSrcFfn
    .TDtaSrcn = TDtaSrcn
End With
End Function
Function TDtaSrc(Fm As TDtaSrcFm, TF() As TF) As TDtaSrc
With TDtaSrc
    .Fm = Fm
    .TF = TF
End With
End Function
Function TFAdd(A As TF, B As TF) As TF(): PushTF TFAdd, A: PushTF TFAdd, B: End Function
Sub PushTFy(O() As TF, A() As TF): Dim J&: For J = 0 To UbTF(A): PushTF O, A(J): Next: End Sub
Sub PushTF(O() As TF, M As TF): Dim N&: N = SiTF(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiTF&(A() As TF): On Error Resume Next: SiTF = UBound(A) + 1: End Function
Function UbTF&(A() As TF): UbTF = SiTF(A) - 1: End Function
Function TF(Tbn, Fny$()) As TF
With TF
    .Tbn = Tbn
    .Fny = Fny
End With
End Function
