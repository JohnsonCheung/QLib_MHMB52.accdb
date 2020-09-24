Attribute VB_Name = "MxXls_Nm"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Nm."
Sub ClrAllWsNm(S As Worksheet): Dim N: For Each N In Itn(S.Names): S.Names(N).Delete: Next: End Sub
Sub ClrAllWbNm(B As Workbook):  Dim N: For Each N In Itn(B.Names): B.Names(N).Delete: Next: End Sub
