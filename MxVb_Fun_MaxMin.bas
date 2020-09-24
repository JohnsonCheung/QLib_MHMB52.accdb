Attribute VB_Name = "MxVb_Fun_MaxMin"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fun_MaxMin."

Function MaxLinCntMdn$()
MaxLinCntMdn = MaxLinCntMdnP(CPj)
End Function

Function MaxLinCntMdnP$(P As VBProject)
MaxLinCntMdnP = Mdn(MaxLinCntMd(P))
End Function
Function MaxLinCntMd(P As VBProject) As CodeModule
Dim C As VBComponent, M&, N&, I
For Each C In P.VBComponents
    N = C.CodeModule.CountOfLines
    If N > M Then
        M = N
        Set MaxLinCntMd = C.CodeModule
    End If
Next
End Function


Function MinUB(Ay1, Ay2)
MinUB = Min(UB(Ay1), UB(Ay2))
End Function

Function MaxVbt(A As VbVarType, B As VbVarType) As VbVarType
Dim O As VbVarType
If A = vbString Or B = vbString Then O = A: Exit Function
If A = vbEmpty Then O = B: Exit Function
If B = vbEmpty Then O = A: Exit Function
If A = B Then O = A: Exit Function
Dim AIsNum As Boolean, BIsNum As Boolean
AIsNum = IsVbtNum(A)
BIsNum = IsVbtNum(B)
Select Case True
Case A = vbBoolean And BIsNum: O = B
Case AIsNum And B = vbBoolean: O = A
Case A = vbDate Or B = vbDate: O = vbString
Case AIsNum And BIsNum:
    Select Case True
    Case A = vbByte: O = B
    Case B = vbByte: O = A
    Case A = vbInteger: O = B
    Case B = vbInteger: O = A
    Case A = vbLong: O = B
    Case B = vbLong: O = A
    Case A = vbSingle: O = B
    Case B = vbSingle: O = A
    Case A = vbDouble: O = B
    Case B = vbDouble: O = A
    Case A = vbCurrency Or B = vbCurrency: O = A
    Case Else: Stop
    End Select
Case Else: Stop
End Select
End Function
