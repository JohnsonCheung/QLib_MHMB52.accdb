Attribute VB_Name = "MxVb_Dta_Itr_ItrRun"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_ItrRun."

Sub ItrDo(Itr, DoFun$):   Dim I: For Each I In Itr: Run DoFun, I: Next: End Sub
Sub ItrDoPX(Itr, PX$, P): Dim I: For Each I In Itr: Run PX, P, I: Next: End Sub
Sub ItrDoXP(Itr, XP$, P): Dim I: For Each I In Itr: Run XP, I, P: Next: End Sub
