Attribute VB_Name = "MxXls_Pt_Op"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Pt_Op."

Sub CpyPt(T As PivotTable, At As Range)
Dim Rg As Range: GoSub Rg
    
Rg.Copy
At.PasteSpecial xlPasteValues
Exit Sub
Rg:
    Dim R1, R2, C1, C2, NC, NR
    Stop
    '    R1 = AA.RowRange.Row
    '    C1 = A.RowRange.Column
    '    R2 = RnoLas(A.DataBodyRange)
    '    C2 = LnoLas(A.DataBodyRange)
        NC = C2 - C1 + 1
        NR = R2 - C1 + 1
    'Set W2RgPt = RgWsRCRC(WsPt(A), R1, C1, R2, C2)
    Return
End Sub
