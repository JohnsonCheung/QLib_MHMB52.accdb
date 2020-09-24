Attribute VB_Name = "MxAcs_Acs_AcsObj_IExpAcs"
Option Compare Text
Option Explicit
Sub ExpAcsC(): ExpAcs Acs, PthSrcPC: End Sub
Sub ExpAcs(A As Access.Application, PthTo$)
Dim P$: P = PthEnsSfx(PthTo)
ChkHasPth P, CSub
ClrPth P
CpyFfnToPth A.CurrentDb.Name, P
ExpAcsSrc A, P
ExpAcsRpt A, P
ExpAcsFrm A, P
ExpAcsQry A, P
ExpAcsRf A, P
End Sub
Sub ImpAcs(A As Access.Application, PthFm$)

End Sub
Sub ExpAcsSrc(A As Access.Application, PthTo$): ExpSrcPthP A.VBE.ActiveVBProject, PthTo: End Sub
Sub ExpAcsRf(A As Access.Application, PthTo$):  ExpRfPthP A.VBE.ActiveVBProject, PthTo:  End Sub
Sub ExpAcsQry(A As Access.Application, PthTo$): ExpQryPth A.CurrentDb, PthTo:            End Sub
