Attribute VB_Name = "MxIde_Md_Op_DltMth1"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_SrcOp_DltMth."
Sub RplMth(M As CodeModule, Mthn, Newl$, Optional ShtMthTy$)
Dim S$(): S = SrcM(M)
Dim Ix&: Ix = Mthix(S, Mthn, ShtMthTy)
Dim NLn%: NLn = NLnMth(S, Ix)
Stop
M.DeleteLines Ix, NLn
M.InsertLines Ix, Newl
End Sub

Sub DltMthny(M As CodeModule, Mthny$())
Dim N: For Each N In Itr(Mthny)
    DltMth M, N
Next
End Sub
Private Sub B_DltMth(): DltMth CMd, "AA": End Sub
Sub DltMth(M As CodeModule, Mthn) ' Dlt mth in Md if exist else inf mth not fnd, if Prp try delete snd time without inf is fnd
Stop '
Dim Bei As Bei
Dim Src$()

Src = SrcM(M)
Bei = BeiMthn(Src, Mthn): If IsEmpBei(Bei) Then Exit Sub
DltCdlLcnt M, LcntBei(Bei)
Debug.Print "DltMth: Mth is DELETED Md[" & M.Parent.Name & "] Mthn[" & Mthn & "]"

Src = SrcM(M)
If Not IsLnPrp(Src(Bei.Bix)) Then Exit Sub

Bei = BeiMthn(Src, Mthn): If IsEmpBei(Bei) Then Debug.Print "DltMth: Mthn is non pair prp, snd Mth is not found. Md[" & M.Parent.Name & "] Mthn[" & Mthn & "]": Exit Sub
DltCdlLcnt M, LcntBei(Bei)
Debug.Print "DltMth: Snd prp mth is DELETED. Md[" & M.Parent.Name & "] Mthn[" & Mthn & "]"
End Sub
Sub DltTMdLnoy(A() As TMdLno)
Dim J&: For J = 0 To UbTMdLno(A)
    With A(J)
        Debug.Print .Md.Parent.Name, .Lno, .Md.Lines(.Lno, 1)
        .Md.DeleteLines .Lno, 1
    End With
Next
End Sub

