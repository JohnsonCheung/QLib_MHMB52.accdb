Attribute VB_Name = "MxDta_Di_CprDi"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Di_CprDi."
Private Type DiCpr
    Kn As String
    H12 As String
    IsExlSam As Boolean
    AExcess As Dictionary
    BExcess As Dictionary
    ADif As Dictionary
    BDif As Dictionary
    Sam As Dictionary
End Type
Public Const SamKeyDifValFf$ = "Key ValA ValB"

Function FmtCprDi(A As Dictionary, B As Dictionary, Optional KeyNm12ss$ = "Key Fst Snd", Optional IsExlSam As Boolean) As String()
FmtCprDi = FmtDiCpr(DiCpr(A, B, KeyNm12ss, IsExlSam))
End Function

Private Sub B_BrwDiCpr()
Dim A As Dictionary, B As Dictionary
Set A = DicVbl("X AA|A BBB|A Lines1|A Line3|B Line1|B line2|B line3..")
BrwDi A
Stop
Set B = DicVbl("X AA|C Line|D Line1|D line2|B Line1|B line2|B line3|B Line4")
BrwDiCpr A, B
End Sub

Sub BrwDiCpr(A As Dictionary, B As Dictionary, Optional NmKeyV12$ = "Key Fst Snd", Optional IsExlSam As Boolean)
BrwAy FmtDiCpr(DiCpr(A, B, NmKeyV12, IsExlSam))
End Sub

Private Function DiKvSam(A As Dictionary, B As Dictionary) As Dictionary
Set DiKvSam = New Dictionary
If A.Count = 0 Or B.Count = 0 Then Exit Function
Dim K
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) = B(K) Then
            DiKvSam.Add K, A(K)
        End If
    End If
Next
End Function

Private Function WFmtDif(A As Dictionary, B As Dictionary, Kn$, H12$) As String()
'@H12:: :NN #Name-1-and-2# Use !AsgT1R to get 2 names
Const CSub$ = CMod & "WFmtDif"
If A.Count <> B.Count Then Thw CSub, "Dic A & B should have same size", "Dic-A-Si Dic-B-Si", A.Count, B.Count
If A.Count = 0 Then Exit Function
Dim O$(), K, S1$, S2$, S() As S12, KK$
For Each K In A
    KK = K
    S1 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & ULnLines(KK) & vbCrLf & A(K)
    S2 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & ULnLines(KK) & vbCrLf & B(K)
    PushS12 S, S12(S1, S2)
Next
WFmtDif = FmtS12y(S, H12:=H12)
End Function

Function DrsKeyV12(A As Dictionary, B As Dictionary, Kn$, Vn1$, Vn2$) As Drs
Const CSub$ = CMod & "DrsKeyV12"
ChkIsDiiVStr A, CSub
ChkIsDiiVStr B, CSub
ChkIsDi12SamKey A, B, CSub
Dim Dy()
    Dim K: For Each K In A.Keys
        PushI Dy, Array(K, A(K), B(K))
    Next
DrsKeyV12 = Drs(Sy(Kn, Vn1, Vn2), Dy)
End Function

Private Function DiCpr(A As Dictionary, B As Dictionary, KeyH12$, IsExlSam As Boolean) As DiCpr
Dim O As DiCpr
Dim L$: L = KeyH12
O.Kn = ShfTm(L)
O.H12 = L
Set O.AExcess = DiMinus(A, B)
Set O.BExcess = DiMinus(B, A)
Set O.Sam = DiKvSam(A, B)
With WDifDi2(A, B)
    Set O.ADif = .A
    Set O.BDif = .B
End With
DiCpr = O
End Function

Private Function WDifDi2(A As Dictionary, B As Dictionary) As Di2
Dim OA As New Dictionary, OB As New Dictionary
Dim K: For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) <> B(K) Then
            OA.Add K, A(K)
            OB.Add K, B(K)
        End If
    End If
Next
WDifDi2 = Di2(OA, OB)
End Function

Private Function FmtDiCpr(A As DiCpr) As String()
Dim O$()
With A
    Dim Nm1$, Nm2$
    AsgT1r A.H12, Nm1, Nm2
    O = AyAddAp( _
        WFmtExcess(.AExcess, .Kn, Nm1), _
        WFmtExcess(.BExcess, .Kn, Nm2), _
        WFmtDif(.ADif, .BDif, .Kn, .H12))
    If Not .IsExlSam Then
        O = AyAdd(O, WFmtSam(A.Sam, .Kn, Nm1, Nm2))
    End If
End With
FmtDiCpr = O
End Function

Private Function WFmtExcess(A As Dictionary, Kn$, Nm$) As String()
If A.Count = 0 Then Exit Function
Dim K, S1$, S2$, S() As S12
For Each K In A.Keys
    S1 = ULnLines(CStr(K))
    S2 = A(K)
    PushS12 S, S12(S1, S2)
Next
PushIAy WFmtExcess, Box(FmtQQ("!Er (?) has Excess", Nm))
PushAy WFmtExcess, FmtS12y(S, H12:="Key " & Nm)
End Function

Private Function WFmtSam(A As Dictionary, Kn$, Nm1$, Nm2$) As String()
If A.Count = 0 Then Exit Function
Dim O$(), K, S() As S12, KK$
For Each K In A.Keys
    KK = K
    PushS12 S, S12("*Same", K & vbCrLf & ULnLines(KK) & vbCrLf & A(K))
Next
WFmtSam = FmtS12y(S)
End Function
