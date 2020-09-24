Attribute VB_Name = "MxVb_Str_Nm_Cml_CCml"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Cml_CCml."

Private Sub B_NmCCmlssy_Ny(): Vc NmCCmlssy_Ny(MthnyPubNonDashPC): End Sub
Function NmCCmlssy_Ny(Ny$()) As String()
Dim N: For Each N In Itr(Ny)
    GoSub Push
Next
Exit Function
Push:
    PushI NmCCmlssy_Ny, CCmlFst(N) & " " & N
    Return
End Function
Function CCmlyFst(Ny$()) As String()
Dim N: For Each N In Itr(Ny)
    PushI CCmlyFst, CCmlFst(N)
Next
End Function
Function CCmlFst$(Nm):       CCmlFst = Left(Nm, LenFstCCml(Nm)): End Function
Function RmvCCmlFst$(Nm): RmvCCmlFst = Mid(Nm, LenFstCCml(Nm)):  End Function

Function CCmly(Nm) As String() '#Camel-Capitial-Array#
Const CSub$ = CMod & "CCmly"
If Nm = "" Then Exit Function
Dim S$: S = Nm
PushI CCmly, ShfCCmlFst(S)
Dim J%
While S <> ""
    ThwLoopTooMuch CSub, J
    PushI CCmly, ShfCCmlRst(S)
Wend
End Function
Private Function ShfCCmlRst$(ONm$) ' fst chr must be UCas
If Not IsAscUCas(Asc(ChrFst(ONm))) Then Thw CSub, "ChrFst of @ONm must be UCas", "@ONm", ONm
ShfCCmlRst = ShfCCmlFst(ONm)
End Function
Function ShfCCmlFst$(ONm$)
Dim L%: L = LenFstCCml(ONm)
If L = 0 Then Exit Function
ShfCCmlFst = Left(ONm, L)
ONm = Mid(ONm, L + 1)
End Function

Function LenFstCCml%(S) 'allow fst chr be UCas
Dim P%
If IsLCas(ChrFst(S)) Then
    P = PosChrUCasFst(S, 2)
    If P = 0 Then
        Thw CSub, "No Cap letter in S", "S", S
    End If
    LenFstCCml = P - 1
    Exit Function
End If
P = PosChrFstNonUCas(S): If P = 0 Then LenFstCCml = Len(S): Exit Function
P = PosChrUCasFst(S, P)
If P = 0 Then
    LenFstCCml = Len(S)
Else
    LenFstCCml = P - 1
End If
End Function

Private Sub B_CCmlyFst(): Vc AySrtQ(AwDis(CCmlyFst(MthnyPC))): End Sub
Private Sub B_CCmllny():  Vc AySrtQ(AwDis(CCmllny(MthnyPC))):  End Sub
Function AetCCml(Ny$()) As Dictionary
Set AetCCml = DiNwSen
Dim N: For Each N In Itr(Ny)
    PushAetAy AetCCml, CCmly(N)
Next
End Function
Function CCmlyDis(Ny$()) As String()
End Function

Function CCmllny(Ny$()) As String()
Dim N: For Each N In Itr(Ny)
    PushI CCmllny, CCmlln(N)
Next
End Function
Function CCmlln$(Nm): CCmlln = Nm & " " & JnSpc(CCmly(Nm)): End Function '#Flat-Camel-SS#

Function Is_ChrLas_UCas(O$) As Boolean
Dim L$: L = ChrLas(O): If L = "" Then Exit Function
Is_ChrLas_UCas = IsAscUCas(Asc(L))
End Function

Function CCmlyy(Ny$()) As String()
Dim N: For Each N In Itr(Ny)
    PushIAy CCmlyy, Cmly(N)
Next
End Function
Function IsAscCmlChr(A%) As Boolean
Select Case True
Case IsAscLetter(A), IsAscDig(A), IsAscDash(A): IsAscCmlChr = True
End Select
End Function

Function IsAscFstcmlchr(A%) As Boolean
If IsAscDash(A) Then Exit Function
IsAscFstcmlchr = IsAscCmlChr(A)
End Function

Function IsUCasCml(Cml$) As Boolean
Select Case True
Case Len(Cml) <> 2, Not IsAscUCas(AscChrFst(Cml)), Not IsAscLCas(AscChrSnd(Cml))
Case Else: IsUCasCml = LCase(ChrFst(Cml)) = ChrSnd(Cml)
End Select
End Function

Function RmvSfxDig$(S)
Dim J%: For J = Len(S) To 1 Step -1
    If Not IsAscDig(Asc(Mid(S, J, 1))) Then RmvSfxDig = Left(S, J): Exit Function
Next
End Function

Function RmvSfxLowDash$(S)
Dim J%: For J = Len(S) To 1 Step -1
    If Mid(S, J, 1) <> "_" Then RmvSfxLowDash = Left(S, J): Exit Function
Next
End Function
