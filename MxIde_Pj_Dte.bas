Attribute VB_Name = "MxIde_Pj_Dte"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Pj_Dte."

Function DteLasMdyFb(Fb) As Date
Dim A As New Access.Application: A.OpenCurrentDatabase Fb
DteLasMdyFb = DteLasMdyAcs(A)
A.CloseCurrentDatabase
End Function

Function DteLasMdyFx(Fx) As Date: DteLasMdyFx = FileDateTime(Fx): End Function

Function DteLasMdyAcs(A As Access.Application)
Dim O As Date
With A.CurrentProject
O = Max(O, MaxItp(.AllForms, "DateModified"))
O = Max(O, MaxItp(.AllModules, "DateModified"))
O = Max(O, MaxItp(.AllReports, "DateModified"))
End With
DteLasMdyAcs = O
End Function
