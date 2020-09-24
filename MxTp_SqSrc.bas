Attribute VB_Name = "MxTp_SqSrc"
Option Compare Text
Option Explicit
Const CMod$ = "MxTp_SqSrc."
Type Stmtll: A As String: End Type
Type Swl: Swn As String: Op As String: Tml() As String: End Type ' Deriving(Ctor Ay)
Type SqSrc
    Sw() As Swl
    Pm() As String
    Stmtlly() As Stmtll
End Type
Function SqSrcT(SqTp$()) As SqSrc
'With SqTpSrcT
'    .Oth = WOth
'    .Pm = WOth
'    .Rmk = Sy()
'    .Sq = WOth
'    .Sw = WOth
'End With
End Function

Private Function WOth() As LLn()

End Function

Function Swl(Swn, Op, Tml$()) As Swl
With Swl
    .Swn = Swn
    .Op = Op
    .Tml = Tml
End With
End Function
Function AddSwl(A As Swl, B As Swl) As Swl(): PushSwl AddSwl, A: PushSwl AddSwl, B: End Function
Sub PushSwlAy(O() As Swl, A() As Swl): Dim J&: For J = 0 To SwlUB(A): PushSwl O, A(J): Next: End Sub
Sub PushSwl(O() As Swl, M As Swl): Dim N&: N = SwlSI(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SwlSI&(A() As Swl): On Error Resume Next: SwlSI = UBound(A) + 1: End Function
Function SwlUB&(A() As Swl): SwlUB = SwlSI(A) - 1: End Function
