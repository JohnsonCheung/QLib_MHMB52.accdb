Attribute VB_Name = "MxIde_Mthn_MthnyUse"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthn_Calln."
Function MthnyUse(Mthy$()) As String()
Dim NyPm$(), NyDim$()
    Dim L$(): L = Contlny(Mthy)
    Dim S$(): S = StmtySrc(CxtMth(L))
    Dim Mthln$: Mthln = L(0)
    Dim Pm$: Pm = BetBkt(Mthln)
    NyDim = VarnyDimy(StmtyWhDim(S))
    NyPm = ArgnyPm(Pm)
    Dim NyUse$(): NyUse = VarnyUse(S)
MthnyUse = SyMinus(SyMinus(NyUse, NyDim), NyPm)
End Function
Function VarnyUse(Stmty$()) As String()
Dim Stmt: For Each Stmt In Itr(Stmty)
    Select Case True
    Case IsStmtDim(Stmt), Stmt = "Next"
    Case IsStmtAsg(Stmt): VarnyUse = VarnyUsezAsg(Stmt)
    Case IsStmtFor(Stmt): VarnyUse = VarnyUsezFor(Stmt)
    End Select
Next
End Function

Function VarnyUsezAsg(Stmt) As String()

End Function
Function VarnyUsezFor(Stmt) As String()

End Function

