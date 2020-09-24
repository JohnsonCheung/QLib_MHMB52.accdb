Attribute VB_Name = "MxIde_Src_Stmt_Dim_GetDimn"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Stmt_Dimn."

Function DimnyStmt(Stmt) As String()
Dim L$: L = Stmt
If IsShfTm(L, "Dim") Then DimnyStmt = AmTrim(SplitCma(LTrim(L)))
End Function

Function DimnyStmty(Stmty$()) As String()
Dim Stmt: For Each Stmt In Itr(Stmty)
    PushI DimnyStmty, DimnyStmt(Stmt)
Next
End Function
