Attribute VB_Name = "MxIde_Src_Msrc1_RplMsrc"
Option Compare Text
Option Explicit
Private Dbg As Boolean
Sub RplMdMsrcDbg(): Dbg = True: RplMdMsrc: Dbg = False: End Sub
Sub RplMdMsrc()
Dim FfnyNew$(): FfnyNew = FfnyMsrcNew
Dim FtNew: For Each FtNew In Itr(FfnyNew)
    RplMdMsrcFtNew FtNew
Next
End Sub
Private Function FfnyMsrcNew() As String(): FfnyMsrcNew = Ffny(PthMsrc, "*(new).txt"): End Function
Private Function PthMsrc$()
Static X$: If X = "" Then X = PthAddFdrEns(PthAssPC, ".Msrc")
PthMsrc = X
End Function
Private Sub RplMdMsrcFtNew(FtMsrcNew)
Dim FtMsrcOld$: FtMsrcOld = FfnRplFnsfx(FtMsrcNew, "(new)", "(old)")
Dim Mdn$: Mdn = Bef(Fn(FtMsrcNew), "("): If Mdn = "MxIde_Src_Msrc_RplMsrc" Then Exit Sub
Dim Newl$: Newl = LinesFt(FtMsrcNew): If Newl = "" Then Raise "Newl is blank": Exit Sub
Dim Oldl$: Oldl = LinesFt(FtMsrcOld): If Oldl = "" Then Raise "Oldl is blank": Exit Sub
Dim Curl$: Curl = SrclMdn(Mdn): If Oldl <> Curl Then MsgBox "Oldl <> Curl": Stop: Exit Sub
If Newl = Oldl Then Raise "Newl = Oldl": Exit Sub
'--
Dim Md As CodeModule: Set Md = CPj.VBComponents(Mdn).CodeModule
Dim NLnOld&: NLnOld = Md.CountOfLines

If Dbg Then
    Debug.Print "RplMdMsrcFtNew: =========================================="
    Debug.Print "Mdn: "; Mdn
    Debug.Print "NLn-New-Old"; NLn(Newl); NLn(Oldl)
    Debug.Print "NLn-New-Len"; Len(Newl)
Else
    Debug.Print FmtQQ("?: Md replaced ? NLn-New-Old ? ? Len-New ?", CSub, AliMax(Mdn, 40), AliR5(NLn(Newl)), AliR5(NLnOld), AliR7(Len(Newl)));
End If
If NLnOld > 0 Then
    If Dbg Then Debug.Print "DbRplMdMsrc: Bef Dlt..."
    Md.DeleteLines 1, NLnOld '<==
    If Dbg Then Debug.Print "DbRplMdMsrc: Bef Ins..."
    Md.InsertLines 1, Newl '<===
End If
If Dbg Then Debug.Print "DbRplMdMsrc: Bef Kill FtMsrcNew & Old..."

Kill FtMsrcNew '<==
Kill FtMsrcOld '<==
End Sub

Private Function Ffny(Pth, Optional Spec$ = "*.*", Optional Atr As VbFileAttribute) As String(): Ffny = AmAddPfx(Fnay(Pth, Spec, Atr), PthEnsSfx(Pth)): End Function
Private Function PthEnsSfx$(Pth)
If Pth = "" Then Exit Function
If HasSfxPth(Pth) Then
    PthEnsSfx = Pth
Else
    PthEnsSfx = Pth & SepPth
End If
End Function
Private Function Fnay(Pth, Optional Spec$ = "*.*", Optional Atr As VbFileAttribute) As String()
Dim O$()
Dim M$: M = Dir(PthEnsSfx(Pth) & Spec, Atr And (Not VbFileAttribute.vbDirectory))
While M <> ""
   PushI Fnay, M
   M = Dir
Wend
End Function

Private Function AmAddPfx(Ay, Pfx) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmAddPfx, Pfx & I
Next
End Function
Private Function PthAddFdrApEns$(Pth, ParamArray FdrAp())
Dim Av(): Av = FdrAp
Dim O$: O = PthAddFdrAv(Pth, Av)
PthEnsAll O
PthAddFdrApEns = O
End Function

Private Function PthAssPC$()
Static P$: If P = "" Then P = PthAssP(CPj)
PthAssPC = P
End Function

Private Function PthAssP$(P As VBProject)
Static A$
Dim B$: B = PthAss(Pjf(P)): If A <> B Then A = PthEns(B)
PthAssP = A
End Function
Private Function CPj() As VBProject: Set CPj = CVbe.ActiveVBProject: End Function

Private Function PthEns$(Pth)
Dim P$: P = PthEnsSfx(Pth)
If NoPth(P) Then MkDir RmvLas(P)
PthEns = P
End Function

Private Function RmvLas$(S): RmvLas = RmvLasN(S, 1): End Function

Private Function RmvLasN$(S, N)
Dim L&: L = Len(S) - N: If L <= 0 Then Exit Function
RmvLasN = Left(S, L)
End Function

Private Function NoPth(Pth) As Boolean
If Not HasPth(Pth) Then Debug.Print "NoPth: "; Pth: NoPth = True
End Function

Private Function HasPth(Pth) As Boolean: HasPth = Fso.FolderExists(Pth): End Function

Private Function Pjf$(P As VBProject)
Pjf = P.FileName
End Function

Private Function CVbe() As VBE: Set CVbe = Application.VBE: End Function

Private Function PthAss$(Ffn): PthAss = Pth(Ffn) & "." & Fn(Ffn) & "\": End Function

Private Function HasSfxPth(Pth) As Boolean: HasSfxPth = ChrLas(Pth) = SepPth: End Function

Private Function Pth$(Ffn)
Dim P%: P = InStrRev(Ffn, "\")
If P = 0 Then Exit Function
Pth = Left(Ffn, P)
End Function

Private Function CutPth$(Ffn)
Dim P%: P = InStrRev(Ffn, SepPth)
If P = 0 Then CutPth = Ffn: Exit Function
CutPth = Mid(Ffn, P + 1)
End Function
Private Function Fn$(Ffn):                          Fn = CutPth(Ffn):                        End Function
Private Function ChrLas$(S):                    ChrLas = Right(S, 1):                        End Function
Private Function PthAddFdrEns$(Pth, Fdr): PthAddFdrEns = PthEns(PthAddFdr(Pth, Fdr)):        End Function
Private Function PthAddFdr$(Pth, Fdr):       PthAddFdr = PthAddSeg(Pth, Fdr):                End Function
Private Function PthAddSeg$(Pth, SegPth):    PthAddSeg = PthEnsSfx(Pth) & PthEnsSfx(SegPth): End Function
Private Sub PushI(O, M)
Dim N&: N = Si(O)
ReDim Preserve O(N)
O(N) = M
End Sub

Private Function Si&(A): On Error Resume Next: Si = UBound(A) + 1: End Function
Private Function UB&(A): UB = Si(A) - 1: End Function

Private Function Itr(Ay)
If Si(Ay) = 0 Then Set Itr = New Collection Else Itr = Ay
End Function

Private Function NLn&(Lines): NLn = Si(SplitCrLf(Lines)): End Function
Private Function SplitCrLf(S) As String()
Dim O$: O = Replace(S, vbCr, "")
SplitCrLf = Split(O, vbLf)
End Function
Private Function FfnRplFnsfx$(Ffn, Sfx$, By$): FfnRplFnsfx = RmvSfx(Ffnn(Ffn), Sfx) & By & Ext(Ffn): End Function
Private Function Ffnn$(Ffn)
Dim B$, C$, P%
B = Fn(Ffn)
P = InStrRev(B, ".")
If P = 0 Then
    C = B
Else
    C = Left(B, P - 1)
End If
Ffnn = Pth(Ffn) & C
End Function
Private Function RmvSfx$(S, Sfx, Optional C As eCas)
If HasSfx(S, Sfx, C) Then RmvSfx = Left(S, Len(S) - Len(Sfx)) Else RmvSfx = S
End Function
Private Function Bef$(S, Ssub$, Optional NoTrim As Boolean, Optional C As eCas)
Dim P%: P = InStr(1, S, Ssub, VbCprMth(C)): If P = 0 Then Exit Function
Bef = Left(S, P - 1)
If Not NoTrim Then Bef = Trim(Bef)
End Function
Private Function VbCprMth(C As eCas) As VbCompareMethod
Const CSub$ = CMod & "CprMth"
Select Case True
Case C = eCasSen: VbCprMth = vbBinaryCompare
Case C = eCasIgn: VbCprMth = vbTextCompare
Case Else: Raise "Invalid value of eCas"
End Select
End Function
Private Function LinesFt$(Ft)
With Fso.GetFile(Ft)
    If .Size = 0 Then Exit Function
    LinesFt = .OpenAsTextStream.ReadAll
End With
End Function
Private Function HasSfx(S, Sfx, Optional C As eCas) As Boolean:  HasSfx = IsEqStr(Right(S, Len(Sfx)), Sfx, C): End Function
Private Function IsEqStr(A, B, Optional C As eCas) As Boolean:  IsEqStr = StrComp(A, B, VbCprMth(C)) = 0:      End Function
Private Function Ext$(Ffn)
Dim B$, P%
B = Fn(Ffn)
P = InStrRev(B, ".")
If P = 0 Then Exit Function
Ext = Mid(B, P)
End Function
Private Sub Raise(Msg$): Err.Raise 1, , Msg: End Sub
Private Function FmtQQ$(QQVbl$, ParamArray Ap())
Dim Av(): Av = Ap
FmtQQ = FmtQQAv(QQVbl, Av)
End Function
Private Function FmtQQAv$(QQVbl$, Av())
Dim O$: O = Replace(QQVbl, "|", vbCrLf)
Dim P&: P = 1
Dim I: For Each I In Av
    P = InStr(P, O, "?")
    If P = 0 Then Exit For
    O = Left(O, P - 1) & Replace(O, "?", I, Start:=P, Count:=1)
    P = P + Len(I)
Next
FmtQQAv = O
End Function
