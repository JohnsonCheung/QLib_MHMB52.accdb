Attribute VB_Name = "MxIde_Dcl_Udt_Udtn"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_Udt_Udtn."



Function UdtnyMC() As String():               UdtnyMC = UdtnyM(CMd):    End Function
Function UdtnyM(M As CodeModule) As String():  UdtnyM = Udtny(DclM(M)): End Function

Function UdtnyPC() As String(): UdtnyPC = UdtnyP(CPj): End Function
Function UdtnyP(P As VBProject) As String():
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy UdtnyP, UdtnyM(C.CodeModule)
Next
End Function

Function Udtny(Dcl$()) As String()
Dim L: For Each L In Itr(Dcl)
    PushNB Udtny, UdtnLn(L)
Next
End Function

Function UdtnyPrvMC() As String():               UdtnyPrvMC = UdtnyPrvM(CMd):    End Function
Function UdtnyPrvM(M As CodeModule) As String():  UdtnyPrvM = UdtnyPrv(DclM(M)): End Function

Function UdtnyPrvPC() As String(): UdtnyPrvPC = UdtnyPrvP(CPj): End Function
Function UdtnyPrvP(P As VBProject) As String():
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy UdtnyPrvP, UdtnyPrvM(C.CodeModule)
Next
End Function
Function UdtnyPrv(Dcl$()) As String()
Dim L: For Each L In Itr(Dcl)
    Dim N$: N = UdtnLn(L)
    If N <> "" Then
        If HasPfx(L, "Private ") Then PushNB UdtnyPrv, N
    End If
Next
End Function

Function IsUdtn(Nm$) As Boolean
Static X$(): If Si(X) = 0 Then X = UdtnyPC
IsUdtn = HasEle(X, Nm)
End Function
