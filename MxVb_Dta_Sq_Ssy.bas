Attribute VB_Name = "MxVb_Dta_Sq_Ssy"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Sq_Ssy."

Function SqSsy(Ssy$()) As Variant(): SqSsy = SqDy(DySsy(Ssy)): End Function
