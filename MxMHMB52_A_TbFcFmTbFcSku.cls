VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_TbFcFmTbFcSku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_FcTb_FmTbFcSku."
Private Type WTSumFcCo
    NFc As Long
    NSku As Long
    SC   As Double
End Type
Private Type WTSumFc
    HK As WTSumFcCo
    MO As WTSumFcCo
End Type
Private Fun As New MxMHMB52_MHFun
Sub RfhLasUD(): RfhStmYm MH.FcLas.StmYmUD: End Sub
Sub RfhLasMH(): RfhStmYm MH.FcLas.StmYmMH: End Sub
Sub RfhStmYm(A As StmYm)
With WTSumFc(A)
RunqC "Update Fc set" & _
" DteLoad=Now," & _
" C86NFc=" & .HK.NFc & "," & _
" C86NSku=" & .HK.NSku & "," & _
" C86Sc=" & .HK.SC & "," & _
" C87NFc=" & .MO.NFc & "," & _
" C87NSku=" & .MO.NSku & "," & _
" C87Sc=" & .MO.SC & _
MH.FcWh.WhStmYm(A)
End With
End Sub

Private Function WTSumFc(A As StmYm) As WTSumFc
Const CSub$ = CMod & "WTSumFc"
Dim W$: W = MH.FcWh.WhStmYm(A)
Dim Sql$: Sql = "Select Distinct Co,Count(*) As NSku," & _
"Sum(Nz(M01,0)+Nz(M02,0)+Nz(M03,0)+Nz(M04,0)+Nz(M05,0)" & _
   "+Nz(M06,0)+Nz(M07,0)+Nz(M08,0)+Nz(M09,0)+Nz(M10,0)" & _
   "+Nz(M11,0)+Nz(M12,0)+Nz(M13,0)+Nz(M14,0)+Nz(M15,0)) As SC," & _
"Sum(" & _
"IIf(Nz(M01,0)=0,0,1)+IIf(Nz(M02,0)=0,0,1)+IIf(Nz(M03,0)=0,0,1)" & _
"+IIf(Nz(M04,0)=0,0,1)+IIf(Nz(M05,0)=0,0,1)+IIf(Nz(M06,0)=0,0,1)" & _
"+IIf(Nz(M07,0)=0,0,1)+IIf(Nz(M08,0)=0,0,1)+IIf(Nz(M09,0)=0,0,1)" & _
"+IIf(Nz(M10,0)=0,0,1)+IIf(Nz(M11,0)=0,0,1)+IIf(Nz(M12,0)=0,0,1)" & _
"+IIf(Nz(M13,0)=0,0,1)+IIf(Nz(M14,0)=0,0,1)+IIf(Nz(M15,0)=0,0,1)) As NFc" & _
" from FcSku" & W & QpGpFf("Co")
Dim Rs As Dao.Recordset: Set Rs = CurrentDb.OpenRecordset(Sql)
With Rs
    While Not .EOF
        Dim Co As Byte: Co = !Co
        Select Case Co
        Case 86: WTSumFc.HK = WTSumFcCo(Rs)
        Case 87: WTSumFc.MO = WTSumFcCo(Rs)
        Case Else: Fun.ThwCo Co, CSub
        End Select
        .MoveNext
    Wend
End With
End Function

Private Function WTSumFcCo(Rs As Recordset) As WTSumFcCo
Dim O As WTSumFcCo
With Rs
    O.NSku = !NSku
    O.SC = !SC
    O.NFc = !NFc
End With
WTSumFcCo = O
End Function
