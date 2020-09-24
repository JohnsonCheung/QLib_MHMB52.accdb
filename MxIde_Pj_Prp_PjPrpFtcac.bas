Attribute VB_Name = "MxIde_Pj_Prp_PjPrpFtcac"
Option Compare Database
Option Explicit

Function MthnyFtcacP(P As VBProject) As String():           MthnyFtcacP = DcStrDy(DyMi8CmntfbelFtcacP(P), 2):     End Function
Function MthnyFtcacPC() As String():                       MthnyFtcacPC = MthnyFtcacP(CPj):                       End Function
Function AetCCmlFtcacP(P As VBProject) As Dictionary: Set AetCCmlFtcacP = AetCCml(MthnyFtcacP(P)):                End Function
Sub VcAetCCmlFtcacP(P As VBProject):                                      VcAet AetCCmlFtcacP(P), "Pj-CCml-Aet ": End Sub
Sub VcAetCCmlFtcacPC():                                                   VcAetCCmlFtcacP CPj:                    End Sub
