Attribute VB_Name = "MxVb_Ay_Op_CvPrimy"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_Cv."

Function CvSy(V) As String():          CvSy = V: End Function
Function CvObj(V) As Object:      Set CvObj = V: End Function
Function CvByty(V) As Byte():        CvByty = V: End Function
Function CvInty(V) As Integer():     CvInty = V: End Function
Function CvBooly(V) As Boolean():   CvBooly = V: End Function
Function CvLngy(V) As Long():        CvLngy = V: End Function
Function CvAv(V) As Variant() 'Ret Av if @V is Av or Empty, else thw error
Const CSub$ = CMod & "CvAv"
Dim T As VbVarType: T = VarType(V)
Select Case True
Case T = vbArray + vbVariant Or T = vbEmpty
    If Si(V) = 0 Then Exit Function
    CvAv = V
    Exit Function
End Select
Thw CSub, "Givan V must be vbArray+vbVariant or vbEmpty", "TypeName(V)", TypeName(V)
End Function

Function CvIntyIf(V) As Integer(): CvIntyIf = ValTrue(IsInty(V), V): End Function
Function CvLngyIf(V) As Long():    CvLngyIf = ValTrue(IsLngy(V), V): End Function


Function CvStr$(V)
Select Case True
Case IsNull(V):
Case IsArray(V): CvStr = "*" & TypeName(V) & "[" & UB(V) & "]"
Case Else: CvStr = V
End Select
End Function
