Attribute VB_Name = "MxIde_Dcl_Udt_TUdt_FmtTUdt"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_Udt_TUdt_FmtTUdt."
Sub A(): BrwTUdtPC: End Sub
Sub BrwTUdtPC(Optional LpmWhTUdt$):                               BrwAy FmtTUdtPC(LpmWhTUdt), "FmtTUdt ": End Sub
Sub VcTUdtPC(Optional LpmWhTUdt$):                                VcAy FmtTUdtPC(LpmWhTUdt), "FmtTUdt ":  End Sub
Function FmtTUdtPC(Optional LpmWhTUdt$) As String():  FmtTUdtPC = FmtTUdty(TUdtyPC(LpmWhTUdt)):           End Function
Function FmtTUdtCmp(C As VBComponent) As String():   FmtTUdtCmp = FmtTUdty(TUdtyCmp(C)):                  End Function
Function FmtTUdt(U As TUdt) As String():                FmtTUdt = FmtTUdty(TUdtySng(U)):                  End Function
Function FmtTUdty(U() As TUdt, Optional LpmWhTUdt$) As String()
Dim UWh() As TUdt
    UWh = TUdtyWh(U, LpmWhTUdt)
    Stop
Dim LMbr$(): LMbr = FmtUdtilnyTUdty(UWh)
Dim LRmk$(): LRmk = FmtUdtyRmk(UWh)
Dim LGen$(): LGen = FmtUdtGen(UWh)
Exit Function
Stop
FmtTUdty = SyAddAp(LMbr, LRmk, LGen)
End Function
Sub B_FmtUdtGen(): BrwAy FmtUdtGen(TUdtyPC): End Sub
Private Function FmtUdtGen(U() As TUdt) As String()
Dim Dy()
    Dim J%: For J = 0 To UbTUdt(U)
        With U(J)
        PushI Dy, Array(.Mdn, .Udtn, _
                StrTrue(.GenCtor, "*Ctor"), _
                StrTrue(.GenAy, "*GenAy"), _
                StrTrue(.GenAdd, "*GenAdd"), _
                StrTrue(.GenPushAy, "*PushAy"), _
                StrTrue(.GenOpt, "*GenOpt"))
        End With
    Next
If Si(Dy) = 0 Then Exit Function
Dim DyHdr(): DyHdr = DyHdrULss("Mdn Udtn Ctor Ay Add PushAy Opt")
Dim Lndy(): Lndy = AyAdd(DyHdr, Dy)
FmtUdtGen = FmtLndy(Lndy, NHdr:=2)
End Function
Private Function FmtUdtyRmk(U() As TUdt) As String()
Dim J%: For J = 0 To UbTUdt(U)
    PushIAy FmtUdtyRmk, FmtUdtRmk(U(J), J + 1)
Next
End Function
Private Function FmtUdtRmk(U As TUdt, NbrUdt%) As String()
PushIAy FmtUdtRmk, FmtUdtRmkHdr(U, NbrUdt)
PushIAy FmtUdtRmk, FmtUdtUmbyRmk(U.Mbr)
Push FmtUdtRmk, ""
End Function
Private Function FmtUdtRmkHdr(U As TUdt, NbrUdt%) As String()
With U
If Si(.Rmky) = 0 Then Exit Function
Dim Dr(): Dr = Array("#" & NbrUdt, IIf(.IsPrv, "Prv", "PUb"), .Mdn & "." & .Udtn, JnCrLf(.Rmky))
FmtUdtRmkHdr = FmtDr(Dr)
End With
End Function
Private Function FmtUdtUmbyRmk(M() As TUmb) As String()
Dim ODy()
Dim J%: For J = 0 To UbTUmb(M)
    PushI ODy, DrUmbRmk(M(J))
Next
FmtUdtUmbyRmk = FmtDy(ODy, eTblFmtTb)
End Function
Private Function DrUmbRmk(M As TUmb) As Variant()
Dim Dr()
PushI Dr, M.Mbn
PushI Dr, M.Tyn & BktIf(M.IsAy)
PushI Dr, JnCrLf(M.Rmky)
DrUmbRmk = Dr
End Function
