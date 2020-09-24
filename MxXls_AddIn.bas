Attribute VB_Name = "MxXls_AddIn"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_AddIn."

Function DrsTAddin(X As Excel.Application) As Drs
DrsTAddin = DrsItp(X.AddIns, "Name FullName Installed IsOpen ProgId CLSID")
End Function
Sub DmpAddinC():                      DmpAddin Xls:        End Sub
Sub DmpAddin(X As Excel.Application): DmpDrs DrsTAddin(X): End Sub

Function WsTAddin(X As Excel.Application) As Worksheet
Set WsTAddin = WsDrs(DrsTAddin(X))
Maxv WsTAddin.Application
End Function

Function Addin(X As Excel.Application, FxaNm) As Excel.Addin
Dim I As Excel.Addin
For Each I In X.AddIns
    If I.Name = FxaNm & ".xlam" Then Set Addin = I
Next
End Function
Function HasAddinFn(B As Excel.Application, FnAddin$) As Boolean: HasAddinFn = HasItn(B.AddIns, FnAddin): End Function
