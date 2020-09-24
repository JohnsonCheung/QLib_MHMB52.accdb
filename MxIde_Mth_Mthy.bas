Attribute VB_Name = "MxIde_Mth_Mthy"
'Cml:Mth :Ly    #Mth-Line-Array# fst ln is Mthln end las ln is MthEln
'Cml:Mthln :Ln  #Mth-Line#       the line with MthTy
'Cml:MthEln :Ln #Mth-End-Line#   the las ln of a Mth
'Cml:Mthy :Lyy  #Array-of-:Mth#  Array of Mth-Line-Array
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_MthMthl."
Function MthyIx(Src$(), Mthix) As String(): MthyIx = AwBE(Src, Mthix, Mtheix(Src, Mthix)): End Function
Function NLnMth%(Src$(), Mthix):            NLnMth = Mtheix(Src, Mthix) - Mthix + 1:       End Function

Private Sub B_MthyySrc():                  VcLyy MthyySrc(SrcPC): End Sub
Function MthyyPC() As Variant(): MthyyPC = MthyySrc(SrcPC):       End Function
Function MthyySrc(Src$()) As Variant()
Dim Ixy&(): Ixy = Mthixy(Src)
Dim Pmpt As Boolean: Pmpt = Si(Ixy) > 1000
If Pmpt Then Debug.Print "MthyySrc: NMth ="; Si(Ixy)
Dim Ix: For Each Ix In Itr(Ixy)
    If Pmpt Then
        Dim J&: J = J + 1
        If J Mod 50 = 0 Then Debug.Print J;
        If J Mod 500 = 0 Then Debug.Print
    End If
    PushI MthyySrc, MthyIx(Src, Ix)
    DoEvents
Next
End Function
Function MthlIx$(Src$(), Mthix): MthlIx = JnCrLf(MthyIx(Src, Mthix)): End Function
Function MthlNmM$(M As CodeModule, Mthn)
Dim S$(): S = SrcM(M)
MthlNmM = MthlIx(S, Mthix(S, Mthn))
End Function

Function CMth() As String(): CMth = SplitCrLf(CMthl): End Function
Function CMthl$()
Dim S$(): S = SrcMC
Dim I&: I = Mthix(S, CMthn)
CMthl = MthlIx(S, I)
End Function

Function CMthln$() 'sdf
Dim M As CodeModule: Set M = CMd
Dim S$(): S = SrcM(M)
Dim I&: I = Mthix(S, CMthn)
CMthln = Mthln(S, I)
End Function
