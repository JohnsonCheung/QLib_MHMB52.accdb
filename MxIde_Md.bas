Attribute VB_Name = "MxIde_Md"
':DiMdDnqSrcl: :Dic ! It is from Pj. Key is Mdn and Val is MdLnes"
':MdDn: :Pjn.Mdn|Mdn
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Md."
Type LLn: Lno As Integer: Ln As String: End Type ' Deriving(Ay Ctor)

Sub ChkMdLno(Fun$, M As CodeModule, Lno&)
If Not IsBet(Lno, 1, M.CountOfLines) Then
    Thw Fun, "Lno is out of Md range", "Mdn Lno Md-Max-Lno", Mdn(M), Lno, M.CountOfLines
End If
End Sub

Function MdDic(P As VBProject, Optional PatnssAndMd$) As Dictionary
Set MdDic = New Dictionary
Dim R As RegExp: Set R = Rx(PatnssAndMd)
Dim C As VBComponent: For Each C In P.VBComponents
    If IsMch(C.Name, R) Then
        MdDic.Add C.Name, SrclM(C.CodeModule)
    End If
Next
End Function
Function MdDicP(Optional PatnssAndMd$) As Dictionary: Set MdDicP = MdDic(CPj, PatnssAndMd):   End Function
Function LnoDclLas&(M As CodeModule):                  LnoDclLas = M.CountOfDeclarationLines: End Function

Function MdCmp(C As VBComponent) As CodeModule
On Error Resume Next
Set MdCmp = C.CodeModule
End Function
Function MdMdn(Mdn) As CodeModule:    Set MdMdn = CmpMdn(Mdn).CodeModule: End Function
Function CmpMdn(Mdn) As VBComponent: Set CmpMdn = CPj.VBComponents(Mdn):  End Function
Function Md(Mddn) As CodeModule
Const CSub$ = CMod & "Md"
Dim A1$(): A1 = Split(Mddn, ".")
Select Case Si(A1)
Case 1: If HasMdn(A1(0)) Then Set Md = CPj.VBComponents(A1(0)).CodeModule
Case 2: If HasMdnP(Pj(A1(0)), A1(1)) Then Set Md = Pj(A1(0)).VBComponents(A1(1)).CodeModule
Case Else: Thw CSub, "[MdDn] should be XXX.XXX or XXX", "MdDn", Mddn
End Select
End Function
Function DiSrclP(P As VBProject) As Dictionary
Set DiSrclP = New Dictionary
Dim C As VBComponent: For Each C In P.VBComponents
    DiSrclP.Add C.Name, SrclM(C.CodeModule)
Next
End Function

Function DiSrclPC() As Dictionary:          Set DiSrclPC = DiSrclP(CPj):               End Function
Function PjnC(A As VBComponent):                    PjnC = A.Collection.Parent.Name:   End Function
Function PjnM(M As CodeModule):                     PjnM = PjnC(M.Parent):             End Function
Function PjM(M As CodeModule) As VBProject:      Set PjM = M.Parent.Collection.Parent: End Function
Function HasPjf(Pjf) As Boolean:                  HasPjf = HasPjfV(CVbe, Pjf):         End Function

Function PjnyXlsC() As String():                      PjnyXlsC = PjnyXls(Xls):                                    End Function
Function PjnyXls(X As Excel.Application) As String():  PjnyXls = PjnyV(X.VBE):                                    End Function
Sub ClsMd(M As CodeModule):                                      M.CodePane.Window.Close:                         End Sub
Sub CprMd(A As CodeModule, B As CodeModule):                     BrwDiCpr DiMthlM(A), DiMthlM(B), Mdn(A), Mdn(B): End Sub
Function CvMd(A) As CodeModule:                       Set CvMd = A:                                               End Function

Private Sub B_HasCdl()
Dim Cdl$: Cdl = "Function HasCdl(M As CodeModule, Cdl$) As Boolea" & "n"
MsgBox HasCdl(CMd, Cdl)
End Sub
Function HasCdl(M As CodeModule, Cdl$) As Boolean:    HasCdl = HasSsub(SrclM(M), Cdl): End Function
Function LinesInfM$(M As CodeModule):              LinesInfM = LinesInf(SrclM(M)):     End Function

Function LLn(Lno, Ln) As LLn
With LLn
    .Lno = Lno
    .Ln = Ln
End With
End Function
Function LLinAdd(A As LLn, B As LLn) As LLn(): PushLLn LLinAdd, A: PushLLn LLinAdd, B: End Function
Sub PushLLnAy(O() As LLn, A() As LLn): Dim J&: For J = 0 To LLnUB(A): PushLLn O, A(J): Next: End Sub
Sub PushLLn(O() As LLn, M As LLn): Dim N&: N = LLnSI(O): ReDim Preserve O(N): O(N) = M: End Sub
Function LLnSI&(A() As LLn): On Error Resume Next: LLnSI = UBound(A) + 1: End Function
Function LLnUB&(A() As LLn): LLnUB = LLnSI(A) - 1: End Function
Function IsMdLE9lines(M As CodeModule) As Boolean
Select Case True
Case M.CountOfLines <= 9: IsMdLE9lines = True
End Select
End Function
Function MdnyLE9linesPC() As String()
Dim C As VBComponent: For Each C In CPj.VBComponents
    If IsMdLE9lines(C.CodeModule) Then PushI MdnyLE9linesPC, C.Name
Next
End Function
