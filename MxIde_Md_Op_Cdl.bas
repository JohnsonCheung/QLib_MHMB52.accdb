Attribute VB_Name = "MxIde_Md_Op_Cdl"
Option Compare Text
Option Explicit
Const TimMdy As Date = #8/24/2020 8:38:11 AM#
Const CMod$ = "MxIde_Md_CdlOp."

Sub RplCdln(M As CodeModule, Lno&, Newln$, Optional Origln$, Optional IsInf As Boolean) ' Replace modul
Const CSub$ = CMod & "RplCdln"
If Origln <> "" Then
    Dim L$: L = M.Lines(Lno, 1)
    If Origln <> L Then Thw CSub, "Md-@Lno-Ln <> Given-@CdlnOrig", "@Lno Md-Lno-Ln @Origln", Lno, L, Origln
End If
M.ReplaceLine Lno, Newln
If IsInf Then
    If L = "" Then L = M.Lines(Lno, 1)
    Inf CSub, "A line is replaced by CdlnNew", "Mdn Lno [CdlnOrig] [CdlnNew]", Mdn(M), Lno, L, Newln
End If
End Sub
Sub EnsCdln(M As CodeModule, Lno&, Cdln$) ' Replace modul
If M.Lines(Lno, 1) = Cdln Then Exit Sub
RplCdln M, Lno, Cdln
End Sub
Sub DltDclln(M As CodeModule, Dclln$)
Dim Lno%: Lno = LnoDclln(M, Dclln)
If Lno > 0 Then M.DeleteLines Lno, 1
End Sub
Sub DltCdl(M As CodeModule, Lno&, Oldl$)
Const CSub$ = CMod & "DltCdl"
If Oldl = "" Then Exit Sub
If Lno = 0 Then Exit Sub
Dim Cnt&: Cnt = NLn(Oldl)
If M.Lines(Lno, Cnt) <> Oldl Then Thw CSub, "OldL <> ActL", "OldL ActL", Oldl, M.Lines(Lno, Cnt)
Debug.Print FmtQQ("DltLines: Lno(?) Cnt(?)", Lno, Cnt)
D Box(SplitCrLf(Oldl))
D ""
M.DeleteLines Lno, Cnt
End Sub

Sub DltCdlCnt(M As CodeModule, Lno&, Optional Cnt = 1, Optional Fun$ = "DltCdlCnt")
Const CSub$ = CMod & "DltCdlCnt"
If Cnt <= 0 Then ThwPm CSub, "Given Cnt should be >=1", "Cnt", Cnt
Dim Cdl$: Cdl = M.Lines(Lno, Cnt)
Inf Fun, "Cdl is deleted.", "Mdn Lno Cnt Cdl", Mdn(M), Lno, Cnt, Cdl
M.DeleteLines Lno, Cnt
End Sub

Sub InsCdl(M As CodeModule, Lno&, Cdl$)
M.InsertLines Lno, Cdl
Debug.Print FmtQQ("InsCdl: Line is ins Lno[?] Md[?] Ln[?]", Lno, Mdn(M), Cdl)
End Sub

Sub DltCdlLcnt(M As CodeModule, A As Lcnt): DltCdlCnt M, A.Lno, A.Cnt: End Sub
Sub AppCdl(M As CodeModule, Cdl$)
If Cdl = "" Then Exit Sub
M.InsertLines M.CountOfLines + 1, Cdl '<=====
End Sub
