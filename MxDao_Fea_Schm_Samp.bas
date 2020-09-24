Attribute VB_Name = "MxDao_Fea_Schm_Samp"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Feat_Schm_Samp."

Function SchmSrcSamp(N%) As SchmSrc: SchmSrcSamp = SchmSrcSchm(SchmSamp(N)):     End Function
Function SchmSamp(N%) As String():      SchmSamp = Resy(X_Fn(N)):                End Function
Sub SchmSampEdt(N%):                               EdtRES X_Fn(N):               End Sub
Private Function X_Fn$(N%):                 X_Fn = "SchmSamp" & N & ".schm.txt": End Function
