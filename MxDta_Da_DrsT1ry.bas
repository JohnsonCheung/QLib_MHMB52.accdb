Attribute VB_Name = "MxDta_Da_DrsT1ry"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Tnr."

Function DrsT1ry(T1ry$(), F12$) As Drs: DrsT1ry = DrsFf(F12, DyTnry(T1ry, 1)): End Function
