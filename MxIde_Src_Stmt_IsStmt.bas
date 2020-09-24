Attribute VB_Name = "MxIde_Src_Stmt_IsStmt"
Option Compare Database
Option Explicit

Function IsStmtWith(Stmt) As Boolean:       IsStmtWith = HasPfxSpc(Stmt, "With"):        End Function
Function IsStmtAsg(Stmt) As Boolean:         IsStmtAsg = HasPfxSpc(Stmt, "With"):        End Function
Function IsStmtWhile(Stmt) As Boolean:     IsStmtWhile = HasPfxSpc(Stmt, "While"):       End Function
Function IsStmtWend(Stmt) As Boolean:       IsStmtWend = Stmt = "Wend":                  End Function
Function IsStmtEndWith(Stmt) As Boolean: IsStmtEndWith = Stmt = "End With":              End Function
Function IsStmtErase(Stmt) As Boolean:     IsStmtErase = HasPfxSpc(Stmt, "Erase"):       End Function
Function IsStmtDo(Stmt) As Boolean:           IsStmtDo = Stmt = "Do":                    End Function
Function IsStmtDim(Stmt) As Boolean:         IsStmtDim = HasPfxSpc(Stmt, "Dim"):         End Function
Function IsStmtReDim(Stmt) As Boolean:     IsStmtReDim = HasPfxSpc(Stmt, "ReDim"):       End Function
Function IsStmtFor(Stmt) As Boolean:         IsStmtFor = HasPfxSpc(Stmt, "For"):         End Function
Function IsStmtIf(Stmt) As Boolean:           IsStmtIf = HasPfxSpc(Stmt, "If"):          End Function
Function IsStmtEndIf(Stmt) As Boolean:     IsStmtEndIf = Stmt = "End If":                End Function
Function IsStmtSelCas(Stmt) As Boolean:   IsStmtSelCas = HasPfxSpc(Stmt, "Select Case"): End Function
Function IsStmtEndSel(Stmt) As Boolean:   IsStmtEndSel = Stmt = "End Select":            End Function
Function IsStmtCas(Stmt) As Boolean:         IsStmtCas = Stmt = "Case":                  End Function
Function IsStmtElse(Stmt) As Boolean:       IsStmtElse = Stmt = "Else":                  End Function
Function IsStmtNxt(Stmt) As Boolean:         IsStmtNxt = HasPfx(Stmt, "Next"):           End Function
