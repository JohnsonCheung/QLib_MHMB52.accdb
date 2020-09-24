Attribute VB_Name = "MxVb_Str_Nm_Nseg"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_A_Nseg."
Function NSegyN(Nm) As String(): Stop ' NSegyNm = SySs(Replace(Nm, "_", " ")): End Function
End Function
Function NSegy_Ny(Ny$()) As String()
Dim O$(): Dim N: For Each N In Itr(Ny)
   Stop ' PushNoDupIAy O, NSegyNm(N)
Next
Stop 'NSegyNy = SySrtQ(O)
End Function
Function NyWhNSegss(Ny$(), WhNSegss$, Optional C As eCas) As String()
Dim WhSegy$(): Stop 'WhSegy = AmAddPfxSfx(SySs(WhSsSeg), "_", "_"): If Si(WhSegy) = 0 Then NyWhSsSeg = SyAy(Ny$()): Exit Function
Dim S: For Each S In Itr(Ny$())
    Stop 'If WHitSegy(S, WhSegy, C) Then PushI NyWhSsSeg, S
Next
End Function
Private Function WHitSegy(S, WhSegy$(), C As eCas) As Boolean
Dim S1$: S1 = "_" & S & "_"
Dim Seg: For Each Seg In WhSegy
    If HasSsub(S1, Seg, C) Then WHitSegy = True: Exit Function
Next
End Function
Function NyWhSsubss(Ay, WhSsubss$, Optional C As eCas) As String()
With Brk1(WhSsubss, ",")
    Dim O$(): O = SyAy(Ay)
    Stop 'O = NyWhSsSeg(O, .S1)
    Stop 'NyWhSsubss = AwSsubss(O, .S2)
End With
End Function

Function NsegyMdnPC() As String(): Stop 'NsegyMdnPC = NSegyNy(MdnyPC):                         End Function
End Function
Sub BrwNsegyMdn(): Vc SySrtQ(NsegyMdnPC), "Nsegy_OfMdn ": End Sub
