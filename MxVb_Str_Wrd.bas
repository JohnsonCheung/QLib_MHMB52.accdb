Attribute VB_Name = "MxVb_Str_Wrd"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Wrd."
Private Type StsLn: Len As Long: NLn As Long: End Type
Private Type StsWrd: Len As Long: NLn As Long: NWrd As Long: End Type
Function FmtCntLn$(Lines):   FmtCntLn = FmtCntLnzSts(StsLn(Lines)):   End Function
Function FmtCntWrd$(Lines): FmtCntWrd = FmtCntWrdzSts(StsWrd(Lines)): End Function
Private Function FmtCntLnzSts$(A As StsLn)
With A
FmtCntLnzSts = FmtQQ("Len ? NLn ?", .Len, .NLn)
End With
End Function
Private Function FmtCntWrdzSts(A As StsWrd)
With A
FmtCntWrdzSts = FmtQQ("Len ? NLn ? NWrd ?", .Len, .NLn, .NWrd)
End With
End Function

Private Function StsWrd(Lines) As StsWrd
With StsWrd
    .NLn = NLn(Lines)
    .Len = Len(Lines)
    .NWrd = NWrd(Lines)
End With
End Function
Private Function StsLn(Lines) As StsLn
With StsLn
    .NLn = NLn(Lines)
    .Len = Len(Lines)
End With
End Function

Function NWrd&(Lines):            NWrd = NMchRx(Lines, RxWrd):     End Function
Function Wrdy(Lines) As String(): Wrdy = SsubyRx(Lines, RxWrd, 1): End Function

Private Function RxWrd() As RegExp
Static X As RegExp: If IsNothing(X) Then Set X = Rx("/[^\d](\w[\w\d_]*)/GIM")
Set RxWrd = X
End Function
