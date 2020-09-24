Attribute VB_Name = "MxIde_Dcl_Udt_TUdt_GetTUdt"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_Udt_TUdt_ToTUdt."

Function TUdtyCmp(C As VBComponent) As TUdt(): TUdtyCmp = TUdtyDcl(DclCmp(C), C.Name): End Function
Function TUdtyMdny(Mdny$()) As TUdt()
Dim N: For Each N In Itr(Mdny)
    PushTUdty TUdtyMdny, TUdtyCmp(CmpMdn(N))
Next
End Function
Function TUdtyDcl(Dcl$(), Optional Mdn$) As TUdt()
Dim Udtyy(): Udtyy = UdtyyDcl(Dcl)
Dim Udty: For Each Udty In Itr(Udtyy)
    PushTUdt TUdtyDcl, TUdtUdty(CvSy(Udty), Mdn)
Next
End Function

Function TUdtyM(M As CodeModule) As TUdt():   TUdtyM = TUdtyDcl(DclM(M), Mdn(M)): End Function
Function TUdtyMC() As TUdt():                TUdtyMC = TUdtyM(CMd):               End Function
Function TUdtyMdn(Mdn) As TUdt():           TUdtyMdn = TUdtyM(Md(Mdn)):           End Function
Function TUdtyPC(Optional LpmWhTUdt$) As TUdt()
Dim P As TLpmBrk: P = TLpmBrk(LpmWhTUdt)
Dim N: For Each N In MdnyPC(Lpmv(P, "Mdn"))
    PushTUdty TUdtyPC, TUdtyM(MdMdn(N))
Next
End Function

Private Sub B_TUdtyM()
GoSub T1
Dim Act() As TUdt, Mdn$
GoSub T2
Exit Sub
T1: Mdn = "MxDaoDbSchmUd": GoTo Tst
T2: Mdn = "MxDaoDbSchmUd": GoTo Tst
Tst:
    Dim M As CodeModule
    Act = TUdtyM(MdMdn(Mdn))
    Stop
    Return
End Sub
Function TUdtMC(Udtn$) As TUdt:                 TUdtMC = TUdtM(CMd, Udtn):                 End Function ' #Fst-Udt-In-CMd#
Function TUdtM(M As CodeModule, Udtn$) As TUdt:  TUdtM = TUdtUdty(UdtyM(M, Udtn), Mdn(M)): End Function
Function TUdtPC(Udtn$) As TUdt '#Fst-Udt-In-CPj#
Dim C As VBComponent: For Each C In CPj.VBComponents
    TUdtPC = TUdtM(C.CodeModule, Udtn)
    If TUdtPC.Udtn <> "" Then Exit Function
Next
End Function
