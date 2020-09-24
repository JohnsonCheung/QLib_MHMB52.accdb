Attribute VB_Name = "MxVb_Dta_Prp_DiPrp"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_DicPrp."
Function StrDivIf$(D As Dictionary, K)
If D.Exists(K) Then StrDivIf = D(K)
End Function
Function DikyStr(D As Dictionary) As String(): DikyStr = SyItr(D.Keys):  End Function ':DikyStr: :Sy ! #Str-Key-Array# it comes from the all keys of a Di
Function Divy(D As Dictionary) As Variant():      Divy = AvItr(D.Items): End Function
Function DivyStr(D As Dictionary, Optional IsSrtKey As Boolean) As String()
If IsSrtKey Then
    DivyStr = DivyStrWhKy(D, AySrtQ(DikyStr(D)))
Else
    DivyStr = SyItr(D.Items)
End If
End Function
Function DivyStrWhKy(D As Dictionary, Ky) As String()
Dim K: For Each K In Itr(Ky)
    PushI DivyStrWhKy, JnCrLf(LinesV(D(K)))
Next
End Function

Function ValDi$(D As Dictionary, K, Optional Dft = Empty)
If D.Exists(K) Then ValDi = D(K) Else ValDi = Dft
End Function

Function ValDiThw$(D As Dictionary, K, Optional Msg$ = "Key not found in Di", Optional Fun$ = "VzDiThw")
If D.Exists(K) Then
    ValDiThw = D(K)
Else
    Thw Fun, Msg, "K DikyStr", K, DikyStr(D)
End If
End Function


Function JnDii(Di As Dictionary, Optional Sep$ = vbCrLf) ' Return the joined Lines from DiLines
JnDii = Jn(Divy(Di), Sep)
End Function

Function StrPfxDi(Pfx$, D As Dictionary) As Dictionary
Dim K: Set StrPfxDi = New Dictionary
For Each K In D.Keys
    StrPfxDi.Add Pfx & K, D(K)
Next
End Function

Function HasBLNKKey(D As Dictionary) As Boolean
If D.Count = 0 Then Exit Function
Dim K: For Each K In D.Keys
   If Trim(K) = "" Then HasBLNKKey = True: Exit Function
Next
End Function

Function HasKy(D As Dictionary, Ky) As Boolean
Ass IsArray(Ky)
If Si(Ky) = 0 Then Stop
Dim K
For Each K In Ky
   If Not D.Exists(K) Then
       Debug.Print FmtQQ("Dix.HasKy: Key(?) is Missing", K)
       Exit Function
   End If
Next
HasKy = True
End Function

Sub HasKyAss(D As Dictionary, Ky)
Dim K
For Each K In Ky
   If Not D.Exists(K) Then Debug.Print K: Stop
Next
End Sub


Private Sub B_IsDikStr()
Dim D As Dictionary
GoSub T1
Exit Sub
T1:
    Set D = New Dictionary
    Dim J&
    For J = 1 To 10000
        D.Add J, J
    Next
    Ept = True
    GoSub Tst
    '
    D.Add 10001, "X"
    Ept = False
    GoTo Tst
Tst:
    Act = IsDikStr(D)
    C
    Return
End Sub

Function TynyAy(Ay) As String():                TynyAy = TynyItr(Itr(Ay)): End Function
Function TynyDii(D As Dictionary) As String(): TynyDii = TynyItr(D.Items): End Function
Function TynyItr(Itr) As String(): Dim V: For Each V In Itr: PushI TynyItr, TypeName(V): Next: End Function

Function VyDiKy(D As Dictionary, Ky) As Variant()
Const CSub$ = CMod & "VyDiKy"
Dim K
For Each K In Itr(Ky)
    If Not D.Exists(K) Then Thw CSub, "Some K in given Ky not found in given Dic keys", "[K with error] [given Ky] [given dic keys]", K, AvItr(D.Keys), Ky
    Push VyDiKy, D(K)
Next
End Function
Function DiWhKy(Di As Dictionary, Ky) As Dictionary
Set DiWhKy = New Dictionary
Dim K: For Each K In Itr(Ky)
    If Di.Exists(K) Then
        DiWhKy.Add K, Di(K)
    End If
Next
End Function

Function DiTy$(D As Dictionary)
Dim O$
Select Case True
Case IsEmpDic(D):   O = "EmpDic"
Case IsDiiStr(D):   O = "StrDic"
Case IsDiiLines(D): O = "LinesDic"
Case IsDiiSy(D):    O = "DiT1qLy"
Case Else:          O = "Dic"
End Select
End Function

Function IsEqDi(A As Dictionary, B As Dictionary) As Boolean
If A.Count <> B.Count Then Exit Function
Dim KeyA: For Each KeyA In A.Keys
    If Not B.Exists(KeyA) Then
        If Not IsEqV(A(KeyA), B(KeyA)) Then Exit Function
    End If
Next
End Function
Function DiLy(LyDi$()) As Dictionary
Set DiLy = New Dictionary
Dim Diln: For Each Diln In Itr(LyDi)
    PushDiln DiLy, CStr(Diln)
Next
End Function
Sub PushDiln(O As Dictionary, Diln$)
With BrkSpc(Diln)
    If O.Exists(.S1) Then
        O(.S1) = O(.S1) & vbCrLf & .S1
    Else
        O.Add .S1, .S2
    End If
End With
End Sub
Function DiAddForce(D As Dictionary, B As Dictionary) As Dictionary
Set DiAddForce = New Dictionary
Set DiAddForce = DiClone(D)
PushDiForce DiAddForce, B
End Function
Function DiAddIf(D As Dictionary, B As Dictionary) As Dictionary
Set DiAddIf = New Dictionary
PushDiForce DiAddIf, D
PushDiForce DiAddIf, B
End Function

Function TynyDiiKy(D As Dictionary, Ky) As String()
Dim K: For Each K In Itr(Ky)
    PushI TynyDiiKy, TypeName(D(K))
Next
End Function
