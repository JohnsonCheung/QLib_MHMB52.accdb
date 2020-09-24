Attribute VB_Name = "MxVb_Dta_Rx_RxChr"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Rxc."
Function RxChrAlpNum() As RegExp
Static X As RegExp: If IsNothing(X) Then Set X = Rx("[0-9a-zA-Z]")
Set RxChrAlpNum = X
End Function
Function RxChrLetter() As RegExp
Static X As RegExp: Set X = Rx("[A-Za-z")
Set RxChrLetter = X
End Function
