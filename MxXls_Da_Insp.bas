Attribute VB_Name = "MxXls_Da_Insp"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Insp."
Enum WdtOpt: eVarWdt: eFixWdt = 1: End Enum
Private Type X
    Wb As Workbook
    IxWs As Worksheet
    IxLo As ListObject
End Type
Private X As X
Sub Init()
WEnsWb
Stop 'EnsIxws B
End Sub
Private Sub WEnsWb()
Stop 'Set X.Wb = WbEns("Insp", Xls): End Sub
End Sub

Sub ClrInsp()
Init
While X.Wb.Sheets.Count > 1
    DltWs X.Wb, 2
Wend
ClrLo X.IxLo
End Sub
Sub InspV(V, Optional N$ = "Var")
Init
Dim Las&
With X.IxLo.ListRows
    .Add
    Las = X.IxLo.ListRows.Count
    Dim S$
    If IsStr(V) Then
        S = "'" & V
    Else
        S = V
    End If
    .Item(Las).Range.Value = SqRowAp(Las, N, Empty, TypeName(V), S, Empty, Empty, Empty)
End With
End Sub

Sub InspDrs(A As Drs, N$, Optional Wdt As WdtOpt = eVarWdt)
Init
Dim Las&, Wsn$, DrsNo%, R As Range
DrsNo = WNxtDrsNo(N)
With X.IxLo.ListRows
    .Add
    Las = X.IxLo.ListRows.Count
    .Item(Las).Range.Value = SqRowAp(Las, N, DrsNo, "Drs", "Go", NRecDrs(A), NDcDrs(A), IsSamDrEleCnt(A))
End With
Wsn = N & DrsNo
Stop 'Set R = DtaDtarg(WsAddDrs(X.Wb, A, Wsn))
If Wdt = WdtOpt.eFixWdt Then
    R.Font.Name = "Courier New"
    R.Font.Size = 9
End If
R.Columns.EntireColumn.AutoFit
Stop
'HLNkypLnk CellLasRow(X.IxLo, "Val"), Wsn
End Sub
Private Function WNxtDrsNo%(DrsNm$)
Dim A As Drs, B As Drs, C As Drs
A = DrsLo(X.IxLo)
Stop 'B = DwEqSel(A, "Nm", DrsNm, "Nm Drs# ValTy")
C = DwEqDrp(B, "ValTy", "Drs")
If NoRecDrs(C) Then WNxtDrsNo = 1: Exit Function
WNxtDrsNo = EleMax(DcIntDrs(C, "Drs#")) + 1
End Function
