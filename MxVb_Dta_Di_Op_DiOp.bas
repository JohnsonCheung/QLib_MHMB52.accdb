Attribute VB_Name = "MxVb_Dta_Di_Op_DiOp"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_DicOp."

Function AddKpfx(A As Dictionary, Kpfx$) As Dictionary
Dim O As New Dictionary
Dim K: For Each K In O.Keys
    O.Add Kpfx & K, A(K)
Next
Set AddKpfx = O
End Function

Function CvDi(A) As Dictionary
Set CvDi = A
End Function

Function AetDik(A As Dictionary) As Dictionary
Set AetDik = AetItr(A.Keys)
End Function

Function DiIup(A As Dictionary, By As Dictionary) As Dictionary 'Return New dictionary from A-Dic by Ins-or-upd By-Dic.  Ins: if By-Dic has key and A-Dic. _
Upd: K fnd in both, A-Dic-Val will be replaced by By-Dic-Val
Dim O As New Dictionary, K
For Each K In A.Keys
    If By.Exists(K) Then
        O.Add K, By(K)
    Else
        O(K) = By(K)
    End If
Next
Set DiIup = O
End Function

Function DiAdd_dickPfx(A As Dictionary, Pfx) As Dictionary
Dim O As New Dictionary, K
For Each K In A.Keys
    O.Add Pfx & K, A(K)
Next
Set DiAdd_dickPfx = O
End Function

Sub DicAddOrUpd(A As Dictionary, K$, V, Sep$)
If A.Exists(K) Then
    A(K) = A(K) & Sep & V
Else
    A.Add K, V
End If
End Sub

Function KyDi(D As Dictionary) As Variant(): KyDi = AyItr(D.Keys): End Function
Function KyDiy(D() As Dictionary) As Variant()
Dim Di: For Each Di In Itr(D)
    Dim K: For Each K In CvDi(Di).Keys
        PushNoDup KyDiy, K
    Next
Next
End Function

Function DiClone(A As Dictionary) As Dictionary
Set DiClone = New Dictionary
Dim K: For Each K In A.Keys
    DiClone.Add K, A(K)
Next
End Function

Function DrDicKy(A As Dictionary, Ky$()) As Variant()
Dim O(), I, J&
ReDim O(UB(Ky))
For Each I In Ky
    If A.Exists(I) Then
        O(J) = A(I)
    End If
    J = J + 1
Next
DrDicKy = O
End Function

Function DyDotlny(Dotlny$()) As Variant()
Dim I, Ln
For Each I In Itr(Dotlny)
    Ln = I
    PushI DyDotlny, SplitDot(Ln)
Next
End Function

Function IntersectDic(A As Dictionary, B As Dictionary) As Dictionary ' ret Sam Key Sam Val
Dim O As New Dictionary
If A.Count = 0 Then GoTo X
If B.Count = 0 Then GoTo X
Dim K
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) = B(K) Then
            O.Add K, A(K)
        End If
    End If
Next
X: Set IntersectDic = O
End Function

Function KeyAet(A As Dictionary) As Dictionary
Set KeyAet = AetItr(A.Keys)
End Function

Function ValDiIfKyJn$(A As Dictionary, Ky, Optional Sep$ = vbCrLf2)
Dim O$(), K
For Each K In Itr(Ky)
    If A.Exists(K) Then
        PushI O, A(K)
    End If
Next
ValDiIfKyJn = Join(O, Sep)
End Function

Function SyDicKy(Dic As Dictionary, Ky$()) As String()
Const CSub$ = CMod & "SyDicKy"
Dim K
For Each K In Itr(Ky)
    If Dic.Exists(K) Then Thw CSub, "K of Ky not in Dic", "K Ky Dic", K, Ky, Dic
    PushI SyDicKy, Dic(K)
Next
End Function

Function LinesDic$(A As Dictionary)
LinesDic = JnCrLf(FmtDiKLines(A))
End Function

Function FmtDiKLines(A As Dictionary) As String()
Dim K: For Each K In A.Keys
    Push FmtDiKLines, LyKLines(K, A(K))
Next
End Function

Function LyKLines(K, Lines$) As String()
Dim Ly$(): Ly = SplitCrLf(Lines)
Dim J&: For J = 0 To UB(Ly)
    Dim Ln
        Ln = Ly(J)
        If ChrFst(Ln) = " " Then Ln = "~" & RmvFst(Ln)
    Push LyKLines, K & " " & Ln
Next
End Function

Function DrsMgeApDi(Hdrnnn$, ParamArray ApDi()) As Drs
Dim AvDi(): AvDi = ApDi
Dim Hdry$(): Hdry = SplitSpc(Hdrnnn)
If Si(AvDi) <> Si(Hdry) Then Stop
Dim AyKey(): AyKey = WAyKey(AvDi)
Dim Dy(): Dy = WDy(AyKey, AvDi)
DrsMgeApDi = Drs(Hdry, Dy)
End Function
Private Function WAyKey(AvDi()) As Variant()
Dim Di, Ix%: For Each Di In AvDi
    Dim K: For Each K In CvDi(AvDi(Di)).Keys
        PushNoDup WAyKey, K
    Next
Next
End Function
Private Function WDy(AyKey(), AvDi()) As Variant()
Dim K: For Each K In AyKey
    PushI WDy, WDr(K, AvDi)
Next
End Function
Private Function WDr(K, AvDi()) As Variant()
Dim U%: U = Si(AvDi)
Dim O(): ReDim O(U)
O(0) = K
Dim Di, Ix%: For Each Di In AvDi
    Ix = Ix + 1
    With CvDi(AvDi(Ix))
        If .Exists(K) Then O(Ix) = .Item(K)
    End With
Next
End Function

Sub PushKvFrc(ODi As Dictionary, K, V)
If Not ODi.Exists(K) Then ODi.Add K, V Else ODi(K) = V
End Sub
Sub PushKvDrp(ODi As Dictionary, K, V)
If Not ODi.Exists(K) Then ODi.Add K, V
End Sub

Function RmvKey(ODi As Dictionary, K) As Boolean
If ODi.Exists(K) Then ODi.Remove K: RmvKey = True
End Function

Function DiMinus(A As Dictionary, B As Dictionary) As Dictionary
'Ret those Ele in A and not in B
If B.Count = 0 Then Set DiMinus = DiClone(A): Exit Function
Dim O As New Dictionary, K
For Each K In A.Keys
   If Not B.Exists(K) Then O.Add K, A(K)
Next
Set DiMinus = O
End Function

Function DicSelIntoAy(A As Dictionary, Ky$()) As Variant()
Dim O()
Dim U&: U = UB(Ky)
ReDim O(U)
Dim J&
For J = 0 To U
   If Not A.Exists(Ky(J)) Then Stop
   O(J) = A(Ky(J))
Next
DicSelIntoAy = O
End Function

Function DicSelIntoSy(A As Dictionary, Ky$()) As String()
DicSelIntoSy = SyAy(DicSelIntoAy(A, Ky))
End Function

Function SwapKv(StrDic As Dictionary) As Dictionary
Set SwapKv = New Dictionary
Dim K: For Each K In StrDic.Keys
    SwapKv.Add StrDic(K), K
Next
End Function

Function WbDiNmqLines(DiNmqLines As Dictionary) As Workbook 'Assume each dic keys is name and each value is lines. _
create a new Wb with worksheet as the dic key and the lines are break to each cell of the sheet
Dim A As Dictionary: Set A = DiNmqLines
Dim K, ThereIsSheet1 As Boolean
Dim O As Workbook: Set O = WbNw
Dim Ws As Worksheet
For Each K In A.Keys
    If K = "Sheet1" Then
        ThereIsSheet1 = True
    Else
        Set Ws = O.Sheets.Add
        Ws.Name = K
    End If
    Ws.Range("A1").Value = SqcLines(A(K))
Next
X: Set WbDiNmqLines = O
End Function

Function DiACoutr(DiAB As Dictionary, DiBC As Dictionary) As Dictionary
Dim A, B, C
Set DiACoutr = New Dictionary
For Each A In DiAB.Keys
    B = DiAB(A)
    If DiBC.Exists(B) Then
        DiACoutr.Add A, C
    Else
        DiACoutr.Add A, Empty
    End If
Next
End Function
Function DiACinr(DiAB As Dictionary, DiBC As Dictionary) As Dictionary
Dim A, B, C
Set DiACinr = New Dictionary
For Each A In DiAB.Keys
    B = DiAB(A)
    If DiBC.Exists(B) Then
        DiACinr.Add A, DiBC(B)
    End If
Next
End Function

Function DiDifVal(A As Dictionary, B As Dictionary) As Dictionary
Set DiDifVal = New Dictionary
Dim K, V
For Each K In A.Keys
    If B.Exists(K) Then
        V = A(K)
        If V <> B(K) Then DiDifVal.Add K, V
    End If
Next
End Function

Sub PushDiIf(O As Dictionary, A As Dictionary)
If IsNothing(O) Then
    Set O = DiClone(A)
    Exit Sub
End If
Dim K
For Each K In A.Keys
    PushKvIf O, K, A(K)
Next
End Sub
Sub PushKv(O As Dictionary, K, V)
Const CSub$ = CMod & "PushKv"
If O.Exists(K) Then Thw CSub, "@K exists in Di-@O", "@K KyDi-@Di", K, KyDi(O)
O.Add K, V
End Sub
Sub PushKvln(O As Dictionary, Kvln)
With BrkSpc(Kvln)
PushKv O, .S1, .S2
End With
End Sub
Sub PushKvIf(ODi As Dictionary, K, V)
If Not ODi.Exists(K) Then ODi.Add K, V
End Sub
Sub PushDi(O As Dictionary, A As Dictionary)
Dim K: For Each K In A.Keys
    PushKv O, K, A(K)
Next
End Sub
Sub PushDiForce(O As Dictionary, A As Dictionary)
Dim K: For Each K In A.Keys
    PushKvForce O, K, A(K)
Next
End Sub
Sub PushKvForce(O As Dictionary, K, V)
If O.Exists(K) Then
    If O(K) <> V Then O(K) = V
Else
    O.Add K, V
End If
End Sub
Function ChainDi(DiAqB As Dictionary, DiBqC As Dictionary) As Dictionary
Const CSub$ = CMod & "ChainDi"
':Ret :DiAqC  ! Thw Er if A->B and B not in fnd in DiBqC @@
Dim A As Dictionary: Set A = DiAqB
Dim B As Dictionary: Set B = DiBqC
Set ChainDi = New Dictionary
Dim KA: For Each KA In A.Keys
    If Not B.Exists(A(KA)) Then
        Thw CSub, "KeyA has ValB which not found in DiBqC", "KeyA ValB DiAqB DiBqC", KA, A(KA), FmtDi(DiAqB), FmtDi(DiBqC)
    End If
    ChainDi.Add KA, B(A(KA))
Next
End Function

Sub PushItmDiT1qLy(A As Dictionary, K, Itm)
Dim M$()
If A.Exists(K) Then
    M = A(K)
    PushI M, Itm
    A(K) = M
Else
    A.Add K, Sy(Itm)
End If
End Sub

Function DiAdd_dicvSfx(D As Dictionary, Sfx$) As Dictionary
Dim O As New Dictionary
Dim K: For Each K In D.Keys
    Dim V$: V = D(K) & Sfx
    O.Add K, V
Next
Set DiAdd_dicvSfx = O
End Function

Sub PushKvNBDrp(ODi As Dictionary, K, StrIfNB$)
If StrIfNB = "" Then Exit Sub
PushKvDrp ODi, K, StrIfNB
End Sub

Function DiSrtStr(D As Dictionary, Optional Ord As eSrt) As Dictionary
If D.Count = 0 Then Set DiSrtStr = New Dictionary: Exit Function
Dim O As New Dictionary
Dim Srt: Srt = SySrtQ(DikyStr(D), Ord, eCasSen)
Dim K: For Each K In Srt
   O.Add K, D(K)
Next
Set DiSrtStr = O
End Function
Function DiSrt(D As Dictionary, Optional Ord As eSrt) As Dictionary
If D.Count = 0 Then Set DiSrt = New Dictionary: Exit Function
Dim O As New Dictionary
Dim Srt: Srt = AySrtQ(D.Keys, Ord)
Dim K: For Each K In Srt
   O.Add K, D(K)
Next
Set DiSrt = O
End Function

Function JnStrVy$(StrDic As Dictionary, Optional Sep$ = vbCrLf2)
JnStrVy = Jn(DivyStr(StrDic), Sep)
End Function

Sub PushDiKqStr(ODiKqStr As Dictionary, K, StrItm, Sep$)
If ODiKqStr.Exists(K) Then
    ODiKqStr(K) = ODiKqStr(K) & Sep & StrItm
Else
    ODiKqStr.Add K, StrItm
End If
End Sub
