Attribute VB_Name = "MxIde_Src_Ens_Mth"
Option Compare Text
Const CMod$ = "MxIde_Src_Ens_Mth."
Option Explicit

Private Sub B_SrcEnsMth()
GoSub T1
Exit Sub
Dim CdlMth$, Src$(), MthnyDlt$()
T1:
    CdlMth = "Private Sub AA(): End Sub" & vbCrLf & "Private Sub BB(): End Sub"
    MthnyDlt = Sy("AA B")
    Src$() = Sy("Private Sub AA():End Sub")
    Ept = SplitCrLf(CdlMth)
    GoTo Tst
Tst:
    Act = SrcEnsMth(Src, CdlMth, MthnyDlt)
    C
    Return
End Sub
Function SrcEnsMth(Src$(), CdlMth$, MthnyDlt$()) As String() '@M should have this @MthCdl, otherwise, dlt @MthnyDlt and ins @MthCdl
Dim Srcl$: Srcl = JnCrLf(Src)
Select Case True
Case CdlMth = ""
    SrcEnsMth = SrcDltMthny(Src, MthnyDlt)
Case HasSsub(Srcl, CdlMth)
    SrcEnsMth = Src
Case Else
    Dim O$(): O = SrcDltMthny(Src, MthnyDlt)
    SrcEnsMth = SyAdd(O, SplitCrLf(CdlMth))
End Select
End Function
Function SrcDltMthny(Src$(), Mthny$()) As String()
Dim O$(): O = Src
Dim Mthn: For Each Mthn In Itr(Mthny)
    O = SrcDltMthn(O, Mthn)
Next
SrcDltMthny = O
End Function
Function SrcDltMthn(Src$(), Mthn) As String(): SrcDltMthn = AeBei(Src, BeiMthn(Src, Mthn)): End Function
