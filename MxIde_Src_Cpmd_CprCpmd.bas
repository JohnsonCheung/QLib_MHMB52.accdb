Attribute VB_Name = "MxIde_Src_Cpmd_CprCpmd"
Option Compare Text
Option Explicit

Sub CprCpmd(A() As Cpmd): CprCpmdl CpmdlyCpmdy(A): End Sub
Sub CprCpmdl(A() As Cpmdl, Optional IsOnlyDif As Boolean)
Dim D() As Cpmdl
If IsOnlyDif Then
    D = CpmdlyWhDif(A)
Else
    D = A
End If
CprLines LinesBef(D), LinesAft(D), "Bef Aft", "Cpr-TMdn-BAft"
End Sub
Private Function LinesBef$(A() As Cpmdl)
Dim OLsy$()
Dim J&: For J = 0 To UbCpmdl(A)
    With A(J)
    Dim Sfx$: Sfx = IIf(.Befl = .Aftl, " (Same)", " (Diff)")
    PushI OLsy, LinesNmLines("#" & J + 1 & " " & .Mdn & Sfx, .Befl) & vbCrLf
    End With
Next
LinesBef = JnCrLf(OLsy)
End Function
Private Function LinesAft$(A() As Cpmdl)
Dim OLsy$()
Dim J&: For J = 0 To UbCpmdl(A)
    With A(J)
    Dim Sfx$: Sfx = IIf(.Befl = .Aftl, " (Same)", " (Diff)")
    PushI OLsy, LinesNmLines("#" & J + 1 & " " & .Mdn & Sfx, .Aftl) & vbCrLf
    End With
Next
LinesAft = JnCrLf(OLsy)
End Function

