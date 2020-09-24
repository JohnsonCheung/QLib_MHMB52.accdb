Attribute VB_Name = "MxTp_Fea_TSpec_Ud"
Option Compare Text
Option Explicit
#If Doc Then
'Udt Spec
'  @Spect snd Tm of fst-line
'  @Specn  third Tm of fst-line
'  @ShtRmk Rst aft third term
'  @Rmk    all indented lines after the fst-ln
'          any dash dash pfx will be removed
'  fst-ln  fst term must be *Spec, otherwise, error.
'Udt TSpeci Spec-Item
'  @Specit Fst Tm of hdr-line
'  @Specin  Snd Tm
'  @ShtRmk  Rst of aft snd term
'  @Rmk     All indent Ly
'  @LLn     following of hdr-line
'  hdr-ln  spec-item-hdr-line.  Non-Identent-Non-DD line
'Cml Catalog#Spec
'  Spec Specification
'  Tp Template
'
'Definition Spec
'  SpecTp
'Cml
' DDSRmk #Hyp-Hyp-Space-Rmk#
#End If
Const CMod$ = "MxTp_Fea_TSpec_Ud."
Type TErSpec: A As String: End Type
Type TIxLn: Ix As Integer: Ln As String: End Type ' Deriving(Ay Ctor)
Type TSpeci: Ix As Integer: Specit As String: Specin As String: Rst As String: IxLny() As TIxLn: End Type 'Deriving(Ay Ctor)
Type TSpec: Spect As String: Specn As String: IndSpec As String: Rmk() As String: Itms() As TSpeci: End Type 'Deriving(Ctor)
Type OptTSpec: Som As Boolean: Spec As TSpec: End Type

Function TIxLn(Ix, Ln) As TIxLn
With TIxLn
    .Ix = Ix
    .Ln = Ln
End With
End Function
Function TIxLnAdd(A As TIxLn, B As TIxLn) As TIxLn(): PushTIxLn TIxLnAdd, A: PushTIxLn TIxLnAdd, B: End Function
Sub PushTIxLny(O() As TIxLn, A() As TIxLn): Dim J&: For J = 0 To UbTIxLn(A): PushTIxLn O, A(J): Next: End Sub
Sub PushTIxLn(O() As TIxLn, M As TIxLn): Dim N&: N = SiTIxLn(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiTIxLn&(A() As TIxLn): On Error Resume Next: SiTIxLn = UBound(A) + 1: End Function
Function UbTIxLn&(A() As TIxLn): UbTIxLn = SiTIxLn(A) - 1: End Function
Function TSpeci(Ix, Specit, Specin, Rst, IxLny() As TIxLn) As TSpeci
With TSpeci
    .Ix = Ix
    .Specit = Specit
    .Specin = Specin
    .Rst = Rst
    .IxLny = IxLny
End With
End Function
Function TSpeciAdd(A As TSpeci, B As TSpeci) As TSpeci(): PushTSpeci TSpeciAdd, A: PushTSpeci TSpeciAdd, B: End Function
Sub PushSpeciy(O() As TSpeci, A() As TSpeci): Dim J&: For J = 0 To UbTSpeci(A): PushTSpeci O, A(J): Next: End Sub
Sub PushTSpeci(O() As TSpeci, M As TSpeci): Dim N&: N = SiTSpeci(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiTSpeci&(A() As TSpeci): On Error Resume Next: SiTSpeci = UBound(A) + 1: End Function
Function UbTSpeci&(A() As TSpeci): UbTSpeci = SiTSpeci(A) - 1: End Function
Function TSpec(Spect, Specn, IndSpec, Rmk$(), Itms() As TSpeci) As TSpec
With TSpec
    .Spect = Spect
    .Specn = Specn
    .IndSpec = IndSpec
    .Rmk = Rmk
    .Itms = Itms
End With
End Function
Function OptTSpec(Som, A As TSpec) As OptTSpec: With OptTSpec: .Som = Som: .Spec = A: End With: End Function
Function SomSpec(A As TSpec) As OptTSpec: SomSpec.Som = True: SomSpec.Spec = A: End Function
