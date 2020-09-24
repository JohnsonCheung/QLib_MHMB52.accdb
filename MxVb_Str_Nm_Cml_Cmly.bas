Attribute VB_Name = "MxVb_Str_Nm_Cml_Cmly"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Cml_Cmly."
Function CmlyFst(Ny$()) As String() ' an array of fst-cml of each @Ny
Dim N: For Each N In Itr(Ny)
    PushI CmlyFst, CmlFst(N)
Next
End Function
Function CmlFst$(Nm): CmlFst = Left(Nm, Wfst_cml_Len(Nm)): End Function

Private Sub B_Cmly()
GoSub Z1
Exit Sub
Dim Nm$
Z1:
    Dim Ny$(): Ny = SySrtQ(AeSfx(MthnyVC, "__Tst"))
    Dim O$(), M$()
    Dim N: For Each N In Ny
        M = Cmly(N)
        If N <> Jn(M) Then Stop
        PushI O, JnSpc(M)
    Next
    Brw FmtT4ry(O)
    Return
T1:
    Nm = "A_IxEle"
    Ept = Sy("A_", "Ix", "Ele")
    GoTo Tst
Tst:
    Act = Cmly(Nm)
    C
    Return
End Sub
Function Cmly(Nm) As String() '#Camel-Capitial-Array#
Const CSub$ = CMod & "Cmly"
If Nm = "" Then Exit Function
Dim S$: S = Nm
PushI Cmly, WShf_CmlFst(S)
Dim J%
While S <> ""
    ThwLoopTooMuch CSub, J
    PushI Cmly, WShfCml(S)
Wend
End Function
Private Function WShfCml$(ONm$)
Const CSub$ = CMod & "WShfCml"
Dim M$: M = WShf_CmlFst(ONm)
If Not IsUCas(M) Then ThwLgc CSub, "Some Non-Fst Cml in @ONm, it not UCas", "@ONm Non-Fst-Cml", ONm, M
WShfCml = M
End Function
Private Function WShf_CmlFst$(ONm$)
Dim L%: L = Wfst_cml_Len(ONm)
If L = 0 Then Exit Function
WShf_CmlFst = Left(ONm, L)
ONm = Mid(ONm, L + 1)
End Function

Private Function Wfst_cml_Len%(Nm)
Dim L%: L = Len(Nm)
Dim J%: For J = 2 To L
    If IsUCas(Mid(Nm, J, 1)) Then
        Wfst_cml_Len = J - 1
        Exit Function
    End If
Next
Wfst_cml_Len = L
End Function

Function CntCmlDis&(Ny$())
Dim O$()
Dim N: For Each N In Itr(Ny)
    PushNoDupIAy O, Cmly(N)
Next
CntCmlDis = Si(O)
End Function
