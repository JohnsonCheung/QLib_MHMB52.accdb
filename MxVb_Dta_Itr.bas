Attribute VB_Name = "MxVb_Dta_Itr"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Itr."
Function MaxItp(Itr, Prpp)
Dim O, Obj: For Each Obj In Itr
    O = Max(O, Opv(Obj, Prpp))
Next
MaxItp = O
End Function

Function NItpTrue(Itr, BoolPrpNm)
Dim O&, X
For Each X In Itr
    If CallByName(X, BoolPrpNm, VbGet) Then
        O = O + 1
    End If
Next
NItpTrue = O
End Function

Function SsItr$(Itr): SsItr = JnSpc(AvItr(Itr)): End Function

Function AvItrMap(Itr, Map$) As Variant(): AvItrMap = IntoItrMap(AvEmp, Itr, Map): End Function

Function NyItr(Itr) As String(): NyItr = Itn(Itr): End Function

Function NyItrEq(Itr, Prpp, V) As String()
Dim Obj: For Each Obj In Itr
    If Opv(Obj, Prpp) = V Then PushI NyItrEq, Objn(Obj)
Next
End Function
Function NyOy(Oy) As String(): NyOy = Itn(Itr(Oy)): End Function

Function VyItrPrpp(Itr, Prpp) As Variant()
Dim Obj: For Each Obj In Itr
    Push VyItrPrpp, Opv(Obj, Prpp)
Next
End Function

Function AyItr(Itr) As Variant()
Dim V: For Each V In Itr
    PushI AyItr, V
Next
End Function
Function Itvy(Itr) As Variant()
Dim I: For Each I In Itr
    PushI Itvy, Objv(I)
Next
End Function
Function Itpv(Itr, Prpp) As Variant()
Dim I: For Each I In Itr
    PushI Itpv, Oppv(I, Prpp)
Next
End Function

Function ItwEq(Itr, Prpp, V) As Variant()
Dim Obj: For Each Obj In Itr
    If Oppv(Obj, Prpp) = V Then Push ItwEq, Obj
Next
End Function

Function Prpny(Itr) As String(): Prpny = Itn(Itr.Properties): End Function
Function ItrLines(Lines$): Asg Itr(SplitCrLf(Lines$)), ItrLines: End Function

Function NItr&(Itr)
Dim O&, V: For Each V In Itr
    O = O + 1
Next
NItr = O
End Function

Private Sub B_Itr()
Dim I
'Set I = Itr(Array(1)) ' This will break
'I = Itr(Array())      ' This will break
Asg Itr(Array()), I         'The will not break
Asg Itr(Array(1)), I        'This will not break
Stop
End Sub
Function ItrTml(Tml$): Asg Itr(Tmy(Tml)), ItrTml: End Function
Function IsItrAllEmp(Itr) As Boolean
Dim I: For Each I In Itr
    If IsEmpty(I) Then Exit Function
Next
IsItrAllEmp = True
End Function

Function Itr(Ay)
If Si(Ay) = 0 Then Set Itr = New Collection Else Itr = Ay
End Function
