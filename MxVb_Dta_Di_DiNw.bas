Attribute VB_Name = "MxVb_Dta_Di_DiNw"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Di_DiFm."
Function DiIxKeEle(Ay) As Dictionary ' ret dic of key is Ele of @Ay and val is Ixy point to that ele
Dim Ixy&(), O As New Dictionary, Ix&
Dim K: For Each K In Itr(Ay)
    If O.Exists(K) Then
        Ixy = O(K)
        PushI Ixy, Ix
        O(K) = Ixy
    Else
        Erase Ixy
        O.Add K, Ixy
    End If
Next
Set DiIxKeEle = O
End Function
Function DiNwIgn() As Dictionary: Set DiNwIgn = DiNw(eCasIgn): End Function
Function DiNwSen() As Dictionary: Set DiNwSen = DiNw(eCasSen): End Function
Function DiNw(C As eCas) As Dictionary
Set DiNw = New Dictionary
DiNw.CompareMode = VbCprMth(C)
End Function
Function DiFt(Ft) As Dictionary
Set DiFt = Diln(LyFt(Ft))
End Function

Function DiTmyKeT1(Tmly$()) As Dictionary
Dim O As New Dictionary
Dim Tml: For Each Tml In Itr(Tmly)
    Dim T$, Rst$
    AsgT1r Tml, T, Rst
    If O.Exists(T) Then
       Stop ' O(T) = AyAdd(O(T), Tml(TmlAy))
    Else
        Stop 'O.Add T, Tml(Rst)
    End If
Next
Set DiTmyKeT1 = O
End Function

Function LyDiLines(DiLines As Dictionary) As String()
Dim Lines$, I
For Each I In DiLines.Items
    Lines = I
    PushIAy LyDiLines, SplitCrLf(Lines)
Next
End Function

Function DiVkkLy(VkkLy$()) As Dictionary
Set DiVkkLy = New Dictionary
Dim I, V$, Vkk$, K
For Each I In Itr(VkkLy)
    Vkk = I
    V = Tm1(Vkk)
    For Each K In SySs(RmvA1T(Vkk))
        DiVkkLy.Add K, V
    Next
Next
End Function

Function LyDi(A As Dictionary, Optional Sep$ = " ") As String()
Dim K
For Each K In A.Keys
    PushI LyDi, K & Sep & A(K)
Next
End Function
Function DiFmDrs(A As Drs, Optional CC$) As Dictionary
If CC = "" Then
    Set DiFmDrs = DiFmDy(A.Dy)
Else
    With BrkSpc(CC)
        Dim C1%: C1 = IxEle(A.Fny, .S1)
        Dim C2%: C2 = IxEle(A.Fny, .S2)
        Set DiFmDrs = DiFmDy(A.Dy, C1, C2)
    End With
End If
End Function

Function DiFmDy(Dy(), Optional C1 = 0, Optional C2 = 1) As Dictionary
Set DiFmDy = New Dictionary
Dim Dr
For Each Dr In Itr(Dy)
    DiFmDy.Add Dr(C1), Dr(C2)
Next
End Function

Function DiLines(Eqlny$(), Optional Sep$ = vbCrLf) As Dictionary 'T1 of each Ly must be Mul
Dim O As New Dictionary, T1$, Rst$
Dim I: For Each I In Itr(Eqlny)
    AsgT1r I, T1, Rst
    If O.Exists(T1) Then
        O(T1) = O(T1) & Sep & Rst
    Else
        O.Add T1, Rst
    End If
Next
Set DiLines = O
End Function
Function Diln(Eqlny$()) As Dictionary 'T1 of each Ly must be uniq otherwise Thw
Const CSub$ = CMod & "Diln"
Set Diln = New Dictionary
Dim T1$, Rst$
Dim I: For Each I In Itr(Eqlny)
    AsgT1r I, T1, Rst
    If Diln.Exists(T1) Then
        Thw CSub, "There are 2 lines with same *Tm1 in @Eqlny", "Er-*TmlAy TRstUnq", T1, Eqlny
    End If
    Diln.Add T1, Rst
Next
End Function

Function DiItmKeAbc(Ay26EleOrLess) As Dictionary
'@Ay26EleOrLess  It is an of *itm with 26 or less ele.  *Itm will be used as the val of the DicVal.
'  And the Ix will 0..25 or less. 0 corrsponding to A,... 25 corresponding to Z
'Ret: Dic with K=A..Z or less and v = corresponding ele in @Ay26orLessEle
Const CSub$ = CMod & "DiItmKeAbc"
If Si(Ay26EleOrLess) > 26 Then Thw CSub, "Si-@Ay26EleOrLess cannot >26", "Si-@Ay26EleOrLess", Si(Ay26EleOrLess)
Dim O As New Dictionary
Dim V, J&: For Each V In Itr(Ay26EleOrLess)
    V = CStr(V)
    If Not O.Exists(V) Then
        O.Add V, Chr(65 + J)
    End If
    J = J + 1
Next
Set DiItmKeAbc = O
End Function

Function DiFmFnyDr(Fny$(), Dr) As Dictionary
Set DiFmFnyDr = New Dictionary
Dim F, J%: For Each F In Fny
    DiFmFnyDr.Add F, Dr(J)
    J = J + 1
Next
End Function

Function DiFmKv(K, V) As Dictionary ' ret a :Di of 1 one element from @K & @v
Set DiFmKv = New Dictionary
DiFmKv.Add K, V
End Function

Function EmpDic() As Dictionary
Set EmpDic = New Dictionary
End Function

Function DiFmKyVy(Ky, Vy) As Dictionary
Const CSub$ = CMod & "DiFmKyVy"
ChkIsEqAySi Ky, Vy, CSub
Dim J&
Set DiFmKyVy = New Dictionary
For J = 0 To UB(Ky)
    DiFmKyVy.Add Ky(J), Vy(J)
Next
End Function

Function DiFmSy12(A$(), B$(), Optional JnSep$ = vbCrLf) As Dictionary
Const CSub$ = CMod & "DiFmSy12"
ChkIsEqAySi A, B, , CSub
Dim O As New Dictionary
Dim I, J&: For Each I In Itr(A)
    If O.Exists(I) Then
       O(I) = O(I) & JnSep & B(J)
    Else
        O.Add I, B(J)
    End If
    J = J + 1
Next
Set DiFmSy12 = O
End Function

Function DiAy12(A, B) As Dictionary
ChkIsEqAySi A, B
Dim N1&, N2&
N1 = Si(A)
N2 = Si(B)
If N1 <> N2 Then Stop
Set DiAy12 = New Dictionary
Dim J&, X
For Each X In Itr(A)
    DiAy12.Add X, B(J)
    J = J + 1
Next
End Function

Function FmtDiEleToIxy(DiEleToIxy As Dictionary) As String()
Dim K: For Each K In DiEleToIxy.Keys
    Dim Ixy&(): Ixy = DiEleToIxy(K)
    PushI FmtDiEleToIxy, StrV(K) & " " & JnSpc(Ixy)
Next
End Function
Function DiEleToIxy(Ay, Optional C As eCas) As Dictionary
Dim O As New Dictionary, Ixy&(), Ix&
Dim V: For Each V In Itr(Ay)
    If O.Exists(V) Then
        Ixy = O(V)
        PushI Ixy, Ix
        O(V) = Ixy
    Else
        ReDim Ixy(0): Ixy(0) = Ix
        O.Add V, Ixy
    End If
    Ix = Ix + 1
Next
Set DiEleToIxy = O
End Function

Function DiCnt(Ay, Optional C As eCas) As Dictionary
':DiCnt: #Cnt-Dic ! Key-Is-Str & Val-Is-CLng
Dim O As New Dictionary
O.CompareMode = VbCprMth(C)
Dim I: For Each I In Itr(Ay)
    If O.Exists(I) Then
        O(I) = O(I) + 1&
    Else
        O.Add I, 1&
    End If
Next
Set DiCnt = O
End Function
