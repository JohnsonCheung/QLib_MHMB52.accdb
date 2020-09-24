Attribute VB_Name = "MxIde_Md_Op_RplMd"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Md_RplMd."

Sub RplMdSrclopt(M As CodeModule, Srclopt As Stropt)
If Srclopt.Som Then RplMd M, Srclopt.Str
End Sub
Sub RplMdSrcopt(M As CodeModule, Srcopt As Lyopt): RplMdSrclopt M, StroptLyopt(Srcopt): End Sub
Private Sub B_RplMd()
Dim M As CodeModule: Set M = Md("QDao_Def_NewTd")
RplMd M, SrclM(M) & vbCrLf & "'"
End Sub
Sub RplMd(M As CodeModule, Newl$)
Const CSub$ = CMod & "RplMd"
If False Then
Select Case M.Name
Case "MxIde_SrcOp_RplLn", "MxVb_Str_Op_Rmv", "MxVb_Run_Thw_Raise", "MxIde_Pj", "MxIde_Mthln_SlmBrw_", "MxVb_Str_Lines", "MxVb_Dta_S12", "MxVb_Run_Thw", "MxIde_Src_Dta"
    Debug.Print "RplMd: " & M.Name, "<== This module is skipped"
    Exit Sub
Case Else
    Debug.Print "RplMd: " & M.Name, "<--- Replacing NLn-" & NLn(Newl)
End Select
End If

Debug.Print "RplMd: " & M.Name, "<--- Replacing NLn-" & NLn(Newl)
If M.CountOfLines > 0 Then
    M.DeleteLines 1, M.CountOfLines '<==
End If
M.InsertLines 1, Newl   '<==
Exit Sub

If Newl = "" Then Thw CSub, "Given Newl is blank"
Dim Oldl$: Oldl = SrclM(M)
Dim IsSam As Boolean: IsSam = LinesEndTrim(Oldl) = LinesEndTrim(Newl)
WShwMsg IsSam, Oldl, Newl, M.Name
If IsSam Then Exit Sub
If M.CountOfLines > 0 Then
    M.DeleteLines 1, M.CountOfLines '<==
End If
M.InsertLines 1, Newl   '<==
End Sub
Private Sub WShwMsg(IsSam As Boolean, Oldl$, Newl$, Mdn$)
Dim Msg$
    Dim NLnO As String * 4: RSet NLnO = NLn(Oldl)
    Dim NLnN As String * 4: RSet NLnN = NLn(Newl)
    Dim MsgRpl As String * 14: MsgRpl = IIf(IsSam, "(Same)", "<=== Replaced")
    Msg = FmtQQ("RplMd: NLn Old/New(?/?) ? ?", NLnO, NLnN, MsgRpl, Mdn)
    Debug.Print Msg
End Sub
