Attribute VB_Name = "MxDta_Da_Dy_Op_DyAli"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Dy_Op_DyAli."

Function DyAli(Dy(), Optional CiyAliR, Optional W% = 150) As Variant() ' ret @@Dy aligned..
'which means (1) each ele is a str. (2) each col is sam wdt.  (3) if the it is lines, each line is same wdt
'@CiyAliR :Int()
If Si(Dy) = 0 Then Exit Function
Dim Wdty%(): Wdty = WdtyDy(Dy, W)
Dim R%(): If IsInty(CiyAliR) Then R = CiyAliR
DyAli = WDy(Dy, Wdty, R)
End Function
Private Function WDy(Dy(), W%(), CiyAliR%()) As Variant()
Dim IsAliRy() As Boolean: IsAliRy = BoolyAliRCiy(CiyAliR, UB(W))
Dim Dr: For Each Dr In Itr(Dy)
    PushI WDy, DrAli(Dr, W, IsAliRy)
Next
End Function

Function DrAliL(Dr, W%()) As String() ' All @Dr.  @W and @IsAliRy are same si.  @Dr is equal or less.
Dim U%: U = UB(W)
Dim O$(): ReDim O(U)
Dim J%: For J = 0 To U
    O(J) = AliL(Dr(J), W(J))
Next
DrAliL = O
End Function

Function DrAli(Dr, W%(), IsAliRy() As Boolean) As String() ' All @Dr.  @W and @IsAliRy are same si.  @Dr is equal or less.
Dim U%: U = UB(W)
Dim O$(): ReDim O(U)
Dim J%: For J = 0 To U
    O(J) = Ali(Dr(J), W(J), IsAliRy(J))
Next
DrAli = O
End Function

Function DrAliQuo$(Dr, W%(), IsAliRy() As Boolean, Q As Qmk): DrAliQuo = LnDr(DrAli(Dr, W, IsAliRy), Q): End Function
