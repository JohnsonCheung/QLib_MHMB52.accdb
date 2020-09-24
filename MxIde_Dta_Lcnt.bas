Attribute VB_Name = "MxIde_Dta_Lcnt"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Udt_Lcnt."
Type Lcnt: Lno As Long: Cnt As Long: End Type 'Deriving(Ctor Ay Opt)
Type Lcnt2: A As Lcnt: B As Lcnt: End Type
Function Lcnt(Lno&, Cnt&) As Lcnt: With Lcnt: .Lno = Lno: .Cnt = Cnt: End With: End Function
Sub PushLcnt(O() As Lcnt, M As Lcnt): Dim N&: N = LcntSI(O): ReDim Preserve O(N): O(N) = M: End Sub
Sub PushLcntAy(O() As Lcnt, A() As Lcnt): Dim J&: For J = 0 To LcntUB(O): PushLcnt O, A(J): Next: End Sub
Function LcntSI&(A() As Lcnt): On Error Resume Next: LcntSI = UBound(A) + 1: End Function
Function LcntUB&(A() As Lcnt): LcntUB = LcntSI(A): End Function
Function LcntStr$(A As Lcnt)
LcntStr = FmtQQ("Lcnt(? ?)", A.Lno, A.Cnt)
End Function
Function Lcnt2(A As Lcnt, B As Lcnt) As Lcnt2: With Lcnt2: .A = A: .B = B: End With: End Function
Function LcntBE(Bix&, Eix&) As Lcnt
With LcntBE
    .Lno = Bix + 1
    .Cnt = Eix - Bix + 1
End With
End Function
Function StrLcnt$(L As Lcnt): StrLcnt = "Lcnt " & L.Lno & " " & L.Cnt: End Function
Function IsEmpLcnt(A As Lcnt) As Boolean
Select Case True
Case A.Cnt <= 0, A.Lno <= 0: IsEmpLcnt = True
End Select
End Function
Function LcntBei(A As Bei) As Lcnt
If IsEmpBei(A) Then Exit Function
LcntBei = Lcnt(A.Bix + 1, A.Eix - A.Bix + 1)
End Function

