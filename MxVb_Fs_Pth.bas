Attribute VB_Name = "MxVb_Fs_Pth"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_Pth."


Function IsEmpPth(Pth) As Boolean
Const CSub$ = CMod & "IsEmpPth"
ChkHasPth Pth, CSub
If HasFfnAny(Pth) Then Exit Function
If HasFdrAny(Pth) Then Exit Function
IsEmpPth = True
End Function

Function PthAddFdrPfx$(Pth, FdrPfx)
With Brk2Rev(PthRmvSfx(Pth), SepPth, NoTrim:=True)
    PthAddFdrPfx = .S1 & SepPth & FdrPfx & .S2 & SepPth
End With
End Function

Function HitFilAtr(A As VbFileAttribute, FilAtr As VbFileAttribute) As Boolean
HitFilAtr = A And FilAtr
End Function

Function FdrFfn$(Ffn): FdrFfn = Fdr(Pth(Ffn)):                  End Function
Function Fdr$(Pth):       Fdr = AftRev(PthRmvSfx(Pth), SepPth): End Function
Function FdrPar$(Pth): FdrPar = Fdr(PthPar(Pth)):               End Function

Sub ChkIsFdrProperNm(Fdr$)
Const CSub$ = CMod & "ChkIsFdrProperNm"
Const C$ = "\/:<>"
If HasChrLis(Fdr, C) Then Thw CSub, "Fdr cannot has these char " & C, "Fdr Char", Fdr, C
End Sub
Function PthPar$(Pth):       PthPar = PthRmvFdr(Pth):         End Function ' Return the PthPar of given Pth
Function PthRmvFdr$(Pth): PthRmvFdr = PthFfn(PthRmvSfx(Pth)): End Function
Function PthUpN$(Pth, UpN%)
Dim O$: O = Pth
Dim J%: For J = 1 To UpN
    O = PthPar(O)
Next
PthUpN = O
End Function

Function PthEns$(Pth)
Dim P$: P = PthEnsSfx(Pth)
If NoPth(P) Then MkDir RmvLas(P)
PthEns = P
End Function
Function FfnEnsPth$(Ffn):   PthEns Pth(Ffn): FfnEnsPth = Ffn: End Function
Function FfnEnsPthAll$(Ffn): PthEnsAll Pth(Ffn): FfnEnsPthAll = Ffn: End Function
Function PthEnsAll$(Pth)
'Ret :Pth and ens each :Pseg. @@
Dim J%, O$, Ay$()
Ay = Split(RmvSfx(Pth, SepPth), SepPth)
O = Ay(0)
For J = 1 To UBound(Ay)
    O = O & SepPth & Ay(J)
    PthEns O
Next
PthEnsAll = PthEnsSfx(Pth)
End Function

Function NoPth(Pth) As Boolean
If Not HasPth(Pth) Then Debug.Print "NoPth: "; Pth: NoPth = True
End Function

Function HasPth(Pth) As Boolean:       HasPth = Fso.FolderExists(Pth):  End Function
Function HasFdr(Pth, Fdr$) As Boolean: HasFdr = HasEle(Fdry(Pth), Fdr): End Function

Sub ChkHasPth(Pth, Fun$):                       ThwTrue NoPth(Pth), Fun, "Pth not exist", "Pth", Pth: End Sub
Function HasFfnAny(Pth) As Boolean: HasFfnAny = Dir(Pth) <> "":                                       End Function
Function HasFdrAny(Pth) As Boolean: HasFdrAny = Fso.GetFolder(Pth).SubFolders.Count > 0:              End Function

Function IsFdryInst(Pth) As String()
Dim Fdr: For Each Fdr In Itr(Fdry(Pth))
    If IsNmInst(Fdr) Then PushI IsFdryInst, Fdr
Next
End Function

Function FdryC(Optional Spec$ = "*.*") As String(): FdryC = Fdry(CDir, Spec): End Function
Function Fdry(Pth, Optional Spec$ = "*.*", Optional Atr As VbFileAttribute) As String()
Dim P$: P = PthEnsSfx(Pth)
Dim E: For Each E In Itr(Enty(P, Spec, Atr))
    If (GetAttr(P & E) And VbFileAttribute.vbDirectory) <> 0 Then
        PushI Fdry, E    '<====
    End If
Next
End Function

Function EntyC(Optional Spec$ = "*.*") As String(): EntyC = Enty(CDir, Spec): End Function
Function Enty(Pth, Optional Spec$ = "*.*", Optional Atr As VbFileAttribute) As String()
Const CSub$ = CMod & "Enty"
ChkHasPth Pth, CSub
Dim A$: A$ = Dir(PthEnsSfx(Pth) & Spec, vbDirectory Or Atr)
While A <> ""
    If A = "." Then GoTo X
    If A = ".." Then GoTo X
    If InStr(A, "?") > 0 Then
        Inf CSub, "Unicode entry is skipped", "UniCode-Entry Pth Spec", A, Pth, Spec
        GoTo X
    End If
    PushI Enty, A
X:
    A = Dir
Wend
End Function

Function FfnItr(Pth): Asg Itr(Ffny(Pth)), FfnItr: End Function
Function Pthy(Pth, Optional Spec$ = "*.*", Optional Atr As VbFileAttribute) As String()
Dim F$(): F = Fdry(Pth, Spec, Atr)
Pthy = AmAddPfxSfx(F, PthEnsSfx(Pth), SepPth)
End Function
Sub ChgCd(Pth)
Const CSub$ = CMod & "ChgCd"
ChkHasPth Pth, "ChgCd"
ChDir Pth
If Not HasPth(Pth) Then Thw CSub, "Pt"
End Sub
Function CDir$():              CDir = CurDir & "\": End Function
Function PthyC() As String(): PthyC = Pthy(CDir):   End Function

Sub AsgEnt(OFdry$(), OFnAy$(), Pth)
Erase OFdry
Erase OFnAy
Dim A$, P$
P = PthEnsSfx(Pth)
A = Dir(Pth, vbDirectory)
While A <> ""
    If A = "." Then GoTo X
    If A = ".." Then GoTo X
    If HasPth(P & A) Then
        PushI OFdry, A
    Else
        PushI OFnAy, A
    End If
    A = Dir
X:
Wend
End Sub

Function Fnny(Pth, Optional Spec$ = "*.*") As String()
Dim I: For Each I In Fnay(Pth, Spec)
    PushI Fnny, Ffnn(I)
Next
End Function

Function FnayFfny(Ffny$()) As String()
Dim I, Ffn$
For Each I In Itr(Ffny)
    Ffn = I
    PushI FnayFfny, Fn(Ffn)
Next
End Function

Function Fnay(Pth, Optional Spec$ = "*.*", Optional Atr As VbFileAttribute) As String()
Dim O$()
Dim M$: M = Dir(PthEnsSfx(Pth) & Spec, Atr And (Not VbFileAttribute.vbDirectory))
While M <> ""
   PushI Fnay, M
   M = Dir
Wend
End Function

Private Sub B_Pthy()
Dim Pth
Pth = "C:\Users\user\AppData\Local\Temp\"
Ept = Sy()
GoSub Tst
Exit Sub
Tst:
    Act = Pthy(Pth)
    Brw Act
    Return
End Sub

Function PthEnsSfx$(Pth)
If Pth = "" Then Exit Function
If HasSfxPth(Pth) Then
    PthEnsSfx = Pth
Else
    PthEnsSfx = Pth & SepPth
End If
End Function
Function SiFfny(Ffny$())
Dim Ffn: For Each Ffn In Itr(Ffny)
    SiFfny = SiFfny + SiFfn(Ffn)
Next
End Function
Function SiFfn&(Ffn)
On Error GoTo X
SiFfn = FileLen(Ffn)
Exit Function
X: Dim E$: E = Err.Description: Infln CSub, E, "Ffn", Ffn
End Function
Private Sub B_Fxy():                                                                                     DmpAy Fxy(CurDir):                              End Sub
Function Fxy(Pth) As String():                                                                     Fxy = Ffny(Pth, "*.xls*"):                            End Function
Function Ffny(Pth, Optional Spec$ = "*.*", Optional Atr As VbFileAttribute) As String():          Ffny = AmAddPfx(Fnay(Pth, Spec, Atr), PthEnsSfx(Pth)): End Function
Function HasSfxPth(Pth) As Boolean:                                                          HasSfxPth = ChrLas(Pth) = SepPth:                           End Function
Function PthRmvSfx$(Pth):                                                                    PthRmvSfx = RmvSfx(Pth, SepPth):                            End Function
Function HasSiblingFdr(Pth, Fdr$) As Boolean:                                            HasSiblingFdr = HasFdr(PthPar(Pth), Fdr):                       End Function
Function SiblingPth$(Pth, SiblingFdr$):                                                     SiblingPth = PthAddFdrEns(PthPar(Pth), SiblingFdr):          End Function
