Attribute VB_Name = "MxXls_Lo_Fmt_Lof_Std"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Lo_Fmt_Lof_Std."

Sub FmtLoWbStd(B As Workbook):  Dim S As Worksheet:  For Each S In B.Sheets:      FmtLoWsStd S: Next: End Sub
Sub FmtLoWsStd(S As Worksheet): Dim L As ListObject: For Each L In S.ListObjects: FmtLoStd L:   Next: End Sub

Sub FmtLoStd(L As ListObject): LoffmtLoUd L, W_LofUd: End Sub
Private Function W_LofUd() As Lofdta
Static X As Lofdta, Y As Boolean: If Not Y Then Y = True: X = LofUd1(X_Lof)
W_LofUd = X
End Function
Private Function X_Lof() As String()

End Function
