Attribute VB_Name = "MxIde_Mth_Sub_Verb"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_Sub_Verb."
Sub BrwVerb(): BrwAy Verby: End Sub
Sub VcVerb():  VcAy Verby:  End Sub
Function Verby() As String()
Static X$(): If Si(X) = 0 Then X = SySs(SsVerb)
Verby = X
End Function
Function AetVerb() As Dictionary
Static X As Dictionary
If IsNothing(X) Then Set X = AetSs(SsVerb)
Set AetVerb = X
End Function

Function SsVerb$()
Const C$ = "Zip Wrt Wrp Wait Vis Vc ULn URmk Unesc Trim Tile Thw Tak Sye Swap Sum Stop" & _
" Srt Split Solve Shw EleShf Set Sel Sav Run Rpl Rmv Rmk Rfh Rev Resi Ren ReSz ReSeq ReOrd RTrim Quo" & _
" Quit Push Pmpt Pop Opn Nxt Nrm New Mov Mk Minus Min Mid Mge Max Map Lnk Lis Lik Las Kill Jn Jmp Is" & _
" Intersect Ins Ini Inf Indt Inc Imp Hit Has Halt Gen Fst Fmt Flat Fill Extend Exp Xls" & _
" Evl Esc Ens EndTrim Edt Drp Down Do Dmp Dlt Cv Cut Crt Cpy Compress Cls Clr Clone Cln" & _
" Chk3 Chk2 Chk1 Chk Chg Brw Bld Below Bef Bdr Bku Aw Ae AutoFit AutoExec Ass Asg" & _
" And Ali Aft Add Above"
SsVerb = SsNrm(C)
End Function
Function SsNrm$(Ss$): SsNrm = JnSpc(AySrtQ(AwDis(SySs(Ss)))): End Function
