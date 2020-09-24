Attribute VB_Name = "MxIde_Src_Dta_TMdLno"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_SrcOp_Dltln_TMdLno."
Type TMdLno: Md As CodeModule: Lno As Long: End Type 'Deriving(Ctor Ay)
Function TMdLno(Md As CodeModule, Lno) As TMdLno
With TMdLno
    Set .Md = Md
    .Lno = Lno
End With
End Function
Sub PushTMdLnoy(O() As TMdLno, A() As TMdLno): Dim J&: For J = 0 To UbTMdLno(A): PushTMdLno O, A(J): Next: End Sub
Sub PushTMdLno(O() As TMdLno, M As TMdLno): Dim N&: N = SiTMdLno(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiTMdLno&(A() As TMdLno): On Error Resume Next: SiTMdLno = UBound(A) + 1: End Function
Function UbTMdLno&(A() As TMdLno): UbTMdLno = SiTMdLno(A) - 1: End Function

Function LyTMdLnoy(A() As TMdLno) As String()
Dim J%: For J = 0 To UbTMdLno(A)
    PushI LyTMdLnoy, LnTMdLno(A(J))
Next
End Function
Private Function LnTMdLno$(A As TMdLno): LnTMdLno = Mdn(A.Md) & " " & A.Md.Lines(A.Lno, 1): End Function

