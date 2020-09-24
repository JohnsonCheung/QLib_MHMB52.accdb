Attribute VB_Name = "MxVb_Dta_Maybe"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Maybe."
Type Lyrslt:  Er() As String: Ly() As String:    End Type
Type Lyopt:   Som As Boolean: Ly() As String:    End Type
Type Syopt:   Som As Boolean: Sy() As String:    End Type
Type Stropt:  Som As Boolean: Str As String:     End Type
Type Boolopt: Som As Boolean: Bool As Boolean:   End Type
Type Diopt:  Som As Boolean: Dic As Dictionary: End Type
Type Lngopt:  Som As Boolean: L As Long:         End Type
Type Intoptt:  Som As Boolean: I As Integer:      End Type
Type Dblopt:  Som As Boolean: D As Double:       End Type

Function Lyrslt(Ly$(), Er$()) As Lyrslt: Lyrslt.Er = Er: Lyrslt.Ly = Ly: End Function
Function SomLng(L) As Lngopt:               SomLng.Som = True:  SomLng.L = L:     End Function
Function SomInt(I) As Intoptt:                 SomInt.Som = True:  SomInt.I = I:     End Function
Function SomSy(Sy$()) As Syopt:               SomSy.Som = True:   SomSy.Sy = Sy:        End Function
Function SomDbl(D) As Dblopt:               SomDbl.Som = True: SomDbl.D = D: End Function
Function SomLy(Ly$()) As Lyopt:               SomLy.Som = True:   SomLy.Ly = Ly:        End Function
Function SomStr(Str) As Stropt:               SomStr.Som = True:  SomStr.Str = Str:     End Function
Function SomBool(Bool As Boolean) As Boolopt: SomBool.Som = True: SomBool.Bool = Bool:  End Function
Function SomDic(Dic As Dictionary) As Diopt: SomDic.Som = True:  Set SomDic.Dic = Dic: End Function
Function SomTrue() As Boolopt:   SomTrue = SomBool(True):  End Function
Function SomFalse() As Boolopt: SomFalse = SomBool(False): End Function
Function StroptLyopt(Lyopt As Lyopt) As Stropt
If Lyopt.Som Then StroptLyopt = SomStr(JnCrLf(Lyopt.Ly))
End Function
Function IsEqStropt(A As Stropt, B As Stropt) As Boolean
Select Case True
Case A.Som And B.Som And A.Str = A.Str: IsEqStropt = True
Case Not A.Som And Not B.Som: IsEqStropt = True
End Select
End Function

Function IsEqLyopt(A As Lyopt, B As Lyopt) As Boolean
Select Case True
Case A.Som And B.Som: IsEqLyopt = IsEqAy(A.Ly, B.Ly)
Case Not A.Som And Not B.Som: IsEqLyopt = True
End Select
End Function

Function LyoptOldNew(LyOld$(), LyNew$()) As Lyopt
If IsEqSy(LyOld, LyNew) Then Exit Function
LyoptOldNew = SomLy(LyNew)
End Function
Function StroptOldNew(StrOld$, StrNew$) As Stropt
If StrOld = StrNew Then Exit Function
StroptOldNew = SomStr(StrNew)
End Function
