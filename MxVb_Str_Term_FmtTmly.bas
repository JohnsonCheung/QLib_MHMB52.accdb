Attribute VB_Name = "MxVb_Str_Term_FmtTmly"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Term_FmtTmly."
Enum eAli: eAliLeft: eAliCentre: eAliRight: End Enum

Sub BrwTmly(Tmly$()): BrwAy FmtTmly(Tmly): End Sub
Sub VcTmly(Tmly$()):  VcAy FmtTmly(Tmly):  End Sub
Sub DmpTmly(Tmly$()): DmpAy FmtTmly(Tmly): End Sub

Function FmtTmly(Tmly$(), Optional CiiAliR$, Optional CiiHypKey$, Optional CiiNbr$, Optional NoIx As Boolean) As String()
If Si(Tmly) = 0 Then Exit Function
FmtTmly = FmtTnry(Tmly, 0, CiiAliR, CiiHypKey, CiiNbr, NoIx)
End Function

Private Sub B_FmtT3ry()
Dim L$()
L = Sy("AAA B C D", "L BBB CCC")
Ept = Sy("AAA B   C   D", _
         "L   BBB CCC")
GoSub Tst
Exit Sub
Tst:
    Act = FmtT3ry(L)
    C
    Return
End Sub
Function FmtT1ry(Tmly$(), Optional CiiAliR$, Optional CiiHypKey$, Optional CiiNbr$, Optional NoIx As Boolean) As String(): FmtT1ry = FmtTnry(Tmly, 1, CiiAliR, CiiHypKey, CiiNbr, NoIx): End Function
Function FmtT2ry(Tmly$(), Optional CiiAliR$, Optional CiiHypKey$, Optional CiiNbr$, Optional NoIx As Boolean) As String(): FmtT2ry = FmtTnry(Tmly, 2, CiiAliR, CiiHypKey, CiiNbr, NoIx): End Function
Function FmtT3ry(Tmly$(), Optional CiiAliR$, Optional CiiHypKey$, Optional CiiNbr$, Optional NoIx As Boolean) As String(): FmtT3ry = FmtTnry(Tmly, 3, CiiAliR, CiiHypKey, CiiNbr, NoIx): End Function
Function FmtT4ry(Tmly$(), Optional CiiAliR$, Optional CiiHypKey$, Optional CiiNbr$, Optional NoIx As Boolean) As String(): FmtT4ry = FmtTnry(Tmly, 4, CiiAliR, CiiHypKey, CiiNbr, NoIx): End Function
Function FmtTnry(Tmly$(), N%, CiiAliR$, Optional CiiHypKey$, Optional CiiNbr$, Optional NoIx As Boolean) As String()
If Si(Tmly) = 0 Then Exit Function
Dim Dy(): Dy = DyTnry(Tmly, N)
FmtTnry = FmtLndy(Dy, CiiAliR, CiiHypKey, CiiNbr, NoIx)
End Function

Private Sub B_FmtT2ry()
Dim L$()
L = Sy("AAA B C D", "L BBB CCC")
Ept = Sy("AAA B   C D", _
         "L   BBB CCC")
GoSub Tst
Exit Sub
Tst:
    Act = FmtT2ry(L)
    C
    Return
End Sub

Function TsyAli(Tsy$()) As String(): TsyAli = FmtDy(DyAli(DyTsy(Tsy))): End Function
Private Function DyTsy(Tsy$()) As Variant()
Dim L: For Each L In Itr(Tsy)
    PushI DyTsy, SplitTab(L)
Next
End Function

Function FmtLySepss(Ly$(), SepSS$) As String(): FmtLySepss = LyDy(DyAli(DyLySepy(Ly, SySs(SepSS)))): End Function
Private Function DyLySepy(Ly$(), Sepy$()) As Variant()
Dim L: For Each L In Itr(Ly)
    PushI DyLySepy, DrLnSepy(L, Sepy)
Next
End Function
Private Function DrLnSepy(Ln, Sepy$(), Optional IsRmvSep As Boolean) As String()
Dim O$()
Dim L$: L = Ln
Dim S: For Each S In Sepy
    PushI O, ShfBef(L, CStr(S))
Next
PushI O, L
If IsRmvSep Then
    Dim J&, Seg: For Each Seg In O
        PushI DrLnSepy, RmvPfx(Seg, Sepy(J))
        J = J + 1
    Next
Else
    DrLnSepy = O
End If
End Function

Function DyTnry(Tmly$(), NTerm%) As Variant() '#Tmly:N-Term-plus-Rest-String-Array#
If NTerm = 0 Then DyTnry = DyTmly(Tmly): Exit Function
Dim L: For Each L In Itr(Tmly)
    PushI DyTnry, DrTnr(L, NTerm)
Next
End Function
Function DrTnr(Tnr, NTerm%) As Variant()
Dim S$: S = Tnr
Dim J%: For J = 1 To NTerm
    PushI DrTnr, ShfTm(S)
Next
PushI DrTnr, S
End Function
Function FmtSsy(Ssy$()) As String(): FmtSsy = LyDy(DyAli(DySsy(Ssy))): End Function

Private Function DySsy(Ssy$()) As Variant()
Dim L: For Each L In Itr(Ssy)
    PushI DySsy, SySs(L)
Next
End Function
