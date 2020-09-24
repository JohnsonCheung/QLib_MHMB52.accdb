Attribute VB_Name = "MxXls_MMinx"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Minx."
Sub Maxv(X As Excel.Application):  WSetWSte X, xlMaximized: ActAppXls X: End Sub
Sub MaxvWs(S As Worksheet):        Maxv S.Application:                   End Sub
Sub MaxvWb(B As Workbook):         Maxv B.Application:                   End Sub
Sub Minv(X As Excel.Application):  WSetWSte X, xlMinimized:              End Sub
Sub Minvn(X As Excel.Application): WSetWSte X, xlNormal, True:           End Sub
Private Sub WSetWSte(X As Excel.Application, S As XlWindowState, Optional IsMinvn As Boolean)
With X
    .Visible = True
    If IsMinvn Then
        .WindowState = xlNormal
        .Top = 1: .Left = 1: .Width = 1: .Height = 1
    Else
        .WindowState = S
    End If
End With
End Sub
