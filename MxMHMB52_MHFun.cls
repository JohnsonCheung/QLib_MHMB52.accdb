VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_MHFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_Fun."
Private Fun As New MxMHMB52_MHFun
Sub ThwCo(Co As Byte, Fun$): Thw Fun, "Co should be 86 or 87, but now[" & Co & "]": End Sub
Sub ChkCo(Co As Byte)
Const CSub$ = CMod & "MhCo_Chk"
If Co <> 86 And Co <> 87 Then ThwCo Co, CSub
End Sub
Function CoStr$(Co As Byte): CoStr = Co & "00": End Function
Function NmCo$(Co As Byte)
Const CSub$ = CMod & "MH.NmCo"
Select Case Co
Case 86: NmCo = "HK"
Case 87: NmCo = "Macau"
Case Else: ThwCo Co, CSub
End Select
End Function
Sub ChkStm(Stm$)
If Not Fun.IsStm(Stm) Then RaiseStm Stm
End Sub
Function IsStm(Stm$) As Boolean
Select Case Stm
Case "U", "M": IsStm = True
End Select
End Function
Function IsStm2(Stm2$) As Boolean
Select Case Stm2
Case "UD", "MH": IsStm2 = True
End Select
End Function

Function StmStm2$(Stm2$)
Dim O$
Select Case Stm2
Case "MH": O = "M"
Case "UD": O = "U"
End Select
StmStm2 = O
End Function
Function StmOth$(Stm$)
Select Case Stm
Case "U": StmOth = "M"
Case "M": StmOth = "U"
Case Else: ThwPm CSub, "Stm should be U or M", "Stm", Stm
End Select
End Function
Function Stm2StmOth$(Stm$): Stm2StmOth = Stm2Stm(StmOth(Stm)): End Function
Function Stm2Stm$(Stm$)
Select Case Stm
Case "U": Stm2Stm = "UD"
Case "M": Stm2Stm = "MH"
Case Else: RaiseStm Stm
End Select
End Function
Function StreamStm$(Stm$)
Select Case Stm
Case "U": StreamStm = "Diageo"
Case "M": StreamStm = "MH"
Case Else: RaiseStm Stm
End Select
End Function
Function StmTopaz$(CdTopaz)
If IsNull(CdTopaz) Then StmTopaz = "?": Exit Function
StmTopaz = IIf(Left(CdTopaz, 3) = "UDV", "U", "M")
End Function
Sub RaiseStm(Stm$): Raise "Stm should be U or M, but now[" & Stm & "]": End Sub
