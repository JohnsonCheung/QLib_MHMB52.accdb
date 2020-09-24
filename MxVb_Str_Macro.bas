Attribute VB_Name = "MxVb_Str_Macro"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Macro."
Function Macrony(Macro$, Optional BktOpn$ = vbBktOpnBig, Optional InlBkt As Boolean) As String()
'Macro is a str with ..[xx].., it is to return all xx or [xx]
Dim Q1$:   Q1 = BktOpn
Dim Q2$:   Q2 = BktCls(BktOpn)
Dim Sy$(): Sy = Split(Macro, Q1)
Dim O$():   O = AwDis(AwNB(BefSyAny(Sy, Q2)))
If InlBkt Then O = AmAddPfxSfx(O, Q1, Q2)
Macrony = O
End Function

Function FmtMacroDi$(Macro$, D As Dictionary)
Dim O$: O = RplVBar(Macro)
Dim K: For Each K In D.Keys
    O = Replace(O, QuoBig(K), D(K))
Next
FmtMacroDi = O
End Function
Function FmtMacroNyAv$(Macro$, Ny$(), Av())
Dim O$: O = Macro
Dim V, J%: For Each V In Av
    O = Replace(O, "{" & Ny(J) & "}", V)
    J = J + 1
Next
FmtMacroNyAv = RplVbl(O)
End Function
Function FmtMacro$(Macro$, NN$, ParamArray Ap())
Dim Av(): Av = Ap
Dim Ny$(): Ny = SySs(NN)
FmtMacro = FmtMacroNyAv(Macro, Ny, Av)
End Function
Function FmtMacroAy$(Macro$, Ay)
Dim Ny$(): Ny = Macrony(Macro)
If Si(Ny) <> Si(Av) Then ThwPm CSub, "Si-@Macrony <> Si-@Ay"
Dim O$: O = RplVBar(Macro)
Dim N, I%: For Each N In Itr(Ny)
    O = Replace(O, "{" & N & "}", Ay(I))
    I = I + 1
Next
FmtMacroAy = O
End Function
Function FmtMacroRs(Macro$, Rs As Dao.Recordset)
FmtMacroRs = FmtMacroDi(Macro, DiRs(Rs))
End Function

Function DiNNAp(NN$, Nav()) As Dictionary
Set DiNNAp = New Dictionary
If Si(Nav) > 0 Then
    Dim Ny$(): Ny = SySs(Nav(0))
    Dim J%: For J = 1 To Si(Ny)
        DiNNAp.Add Ny(J - 1), Nav(J)
    Next
End If
End Function
