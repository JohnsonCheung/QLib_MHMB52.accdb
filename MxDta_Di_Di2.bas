Attribute VB_Name = "MxDta_Di_Di2"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Di_Di2."
Type Di2
    A As Dictionary
    B As Dictionary
End Type
Function Di2(A As Dictionary, B As Dictionary) As Di2
Const CSub$ = CMod & "Di2"
ChkIsVSomg A, "DicA", CSub
ChkIsVSomg B, "DicB", CSub
With Di2
    Set .A = A
    Set .B = B
End With
End Function
Function Di2InKy(D As Dictionary, InKy) As Di2
Dim K, A As New Dictionary, B As New Dictionary
For Each K In D.Keys
    If HasEle(InKy, K) Then
        A.Add K, D(K)
    Else
        B.Add K, D(K)
    End If
Next
Di2InKy = Di2(A, B)
End Function
