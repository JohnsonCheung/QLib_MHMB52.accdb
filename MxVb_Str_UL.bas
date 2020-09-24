Attribute VB_Name = "MxVb_Str_UL"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_UL."

Function ULnHdr$(Hdr$)
Dim L%: L = Len(Hdr)
Dim O$: O = Space(L)
Dim J%: For J = 1 To L
    If Mid(Hdr, J, 1) <> " " Then Mid(O, J, 1) = "="
Next
ULnHdr = O
End Function

Function ULnLines$(Lines$, Optional ChrUL$ = "="): ULnLines = String(WdtLines(Lines), ChrUL):          End Function
Function LinesUL$(Lines$, Optional ChrUL$ = "="):   LinesUL = Lines & vbCrLf & ULnLines(Lines, ChrUL): End Function
Function LyUL(S, Optional ChrUL$ = "=") As String()
PushI LyUL, S
PushI LyUL, ULn(S, ChrUL)
End Function
Function ULn$(L, Optional ChrUL$ = "="): ULn = String(Len(L), ChrUL): End Function
Function DyHdrULss(HdrULss$, Optional ChrUL$ = "=") As Variant()
Dim H$()
Dim UL$()
    H = SySs(HdrULss)
    Dim I: For Each I In H
        PushI UL, String(Len(I), ChrUL$)
    Next
DyHdrULss = Array(H, UL)
End Function
Sub PushLyUL(O$(), S$, Optional ChrUL$ = "="): PushIAy O, LyUL(S, ChrUL): End Sub
