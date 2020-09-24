Attribute VB_Name = "MxAcs_Acs_AcsObj_CvAcsObj"
Option Compare Text
Option Explicit
Const CMod$ = "MxAcs_Acs_CvAcsObj."
Function CvBtn(A) As Access.CommandButton: Set CvBtn = A: End Function
Function CvCtl(A) As Access.Control:       Set CvCtl = A: End Function
Function CvTgl(A) As Access.ToggleButton:  Set CvTgl = A: End Function
Function CvAcs(A) As Access.Application:   Set CvAcs = A: End Function
