VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_IsLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_IsLoad."
Function Git(A As Ymd)
Git = HasRecQC("Select * from Git" & MH.TbOH.WhYmd(A))
If Git Then Exit Function
Git = MsgBox("Git is not loaded!!" & vbCrLf & "[Ok] = Contine generate report" & vbCrLf & "[Cancel] = Cancel", vbOKCancel + vbQuestion) = vbCancel
End Function
Function MB52(A As Ymd) As Boolean
MB52 = HasRecQC("Select Top 1 Count(*) from OH" & MH.TbOH.WhYmd(A))
If MB52 Then Exit Function
MsgBox "Please [Load MB52] first", vbCritical
End Function
