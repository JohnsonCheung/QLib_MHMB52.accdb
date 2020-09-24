Attribute VB_Name = "MxXls_Wb_Op"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Wb_Op."

Sub SavWbQuit(Wb As Workbook)
Dim X As Excel.Application
Set X = Wb.Application
Wb.Close True
X.Quit
End Sub

Sub ArrangeWbV(X As Excel.Application)
Dim W As Excel.Window: For Each W In X.Windows
    W.Activate
    W.WindowState = xlNormal
    W.Visible = True
Next
X.Windows.Arrange xlArrangeStyleVertical
End Sub
