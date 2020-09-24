Attribute VB_Name = "MxVb_Dta_Lpm_WhNm"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_WhNm."
Type TWhNm
    Rx As RegExp
    ExlLiky() As String
    Liky() As String
    IsEmp As Boolean
End Type

Function TWhNmEmp() As TWhNm: TWhNmEmp.IsEmp = True: End Function

Function IsEqTWhNm(A As TWhNm, B As TWhNm) As Boolean
With A
    If .IsEmp And B.IsEmp Then IsEqTWhNm = True: Exit Function
Select Case True
Case _
    ObjPtr(.Rx) <> ObjPtr(.Rx), _
    IsEqAy(.ExlLiky, B.ExlLiky), _
    IsEqAy(.Liky, B.Liky)
Case Else
    IsEqTWhNm = True
End Select
End With
End Function


Function FmtWhNm$(A As TWhNm)
'If IsEmpWhNm(A) Then ToStr = "#Emp": Exit Function
Dim O$()
'Push O, Quo(X_Re.Pattern, "Patn(*)")
'Push O, Quo(TLn(X_Liky), "Liky(*)")
'Push O, Quo(TLn(X_ExlLiky), "ExlLiky(*)")
'ToStr = JnCrLf(O)
End Function

Function IsEmpTWhNm(A As TWhNm) As Boolean: IsEmpTWhNm = A.IsEmp: End Function
Function HitLpmWhNm(S, LpmWhNm$) As Boolean
HitLpmWhNm = HitWhNm(S, TWhNmLpm(LpmWhNm))
End Function
Function TWhNmLpm(LpmWhNm$) As TWhNm
Stop
End Function
Function HitWhNm(S, A As TWhNm) As Boolean
HitWhNm = True
With A
If .IsEmp Then Exit Function
If HitLiky(S, .ExlLiky) Then HitWhNm = False: Exit Function
If HitRx(S, .Rx) Then Exit Function
If HitLiky(S, .Liky) Then Exit Function
End With
HitWhNm = False
End Function
Function AwWhNm(Ay, B As TWhNm) As String()
Dim I: For Each I In Itr(Ay)
    If HitWhNm(I, B) Then PushI AwWhNm, I
Next
End Function

Function AwLpmWhNm(Ay, LpmWhNm$) As String(): AwLpmWhNm = AwWhNm(Ay, TWhNmLpm(LpmWhNm)): End Function

