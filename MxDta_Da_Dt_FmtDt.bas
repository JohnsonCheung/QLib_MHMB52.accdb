Attribute VB_Name = "MxDta_Da_Dt_FmtDt"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Dt_FmtDt."
Function FmtDs(D As Ds, _
Optional Rds As eRds, _
Optional Fmt As eTblFmt, _
Optional Zer As eZer, _
Optional ColnnSum$, _
Optional Wdt% = 100) As String()
PushI FmtDs, "*Ds " & D.Dsn & " " & String(10, "=")
Dim J%: For J = 0 To DtUB(D.Dty)
    PushAy FmtDs, FmtDt(D.Dty(J), Rds, Fmt, Zer, ColnnSum, Wdt)
Next
End Function
Function FmtDt(D As Dt, _
Optional Rds As eRds, _
Optional Fmt As eTblFmt, _
Optional Zer As eZer, _
Optional ColnnSum$, _
Optional Wdt% = 100) As String()
End Function
Sub BrwDt(Dt As Dt, _
Optional Rds As eRds, _
Optional Fmt As eTblFmt, _
Optional Zer As eZer, _
Optional ColnnSum$, _
Optional Wdt% = 100, _
Optional PfxFn$ = "Ds_")
Dim F$(): F = FmtDt(Dt, Rds, Fmt, Zer, ColnnSum, Wdt)
BrwAy F, PfxFn
End Sub

Sub BrwDs(D As Ds, _
Optional Rds As eRds, _
Optional Fmt As eTblFmt, _
Optional Zer As eZer, _
Optional ColnnSum$, _
Optional Wdt% = 100, _
Optional PfxFn$ = "Ds_")
BrwAy FmtDs(D, Rds, Fmt, Zer, ColnnSum, Wdt), PfxFn
End Sub

Sub DmpDs(D As Ds, _
Optional Rds As eRds, _
Optional Fmt As eTblFmt, _
Optional Zer As eZer, _
Optional ColnnSum$, _
Optional Wdt% = 100)
DmpAy FmtDs(D, Rds, Fmt, Zer, ColnnSum, Wdt)
End Sub
