Attribute VB_Name = "MxIde_Pj_Bku"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Pj_Bku."

Sub BkuPjC(Optional Msg$ = "Bku"): BkuPj CPj, Msg: End Sub

Sub BkuPj(P As VBProject, Optional Msg$ = "Bku"):             BkuFfn Pjf(P), Msg:  End Sub
Sub BrwPthBku():                                              BrwPth PthBkuP(CPj): End Sub
Function PthBkuP$(P As VBProject):                  PthBkuP = PthBku(Pjf(P)):      End Function
Function FfnyBkuPC() As String():                 FfnyBkuPC = FfnyBku(CPjf):       End Function
Function PjfBkuLas$():                            PjfBkuLas = FfnBkuLas(CPjf):     End Function
