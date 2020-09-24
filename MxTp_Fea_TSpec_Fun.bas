Attribute VB_Name = "MxTp_Fea_TSpec_Fun"
Option Compare Text
Option Explicit
Const CMod$ = "MxTp_Fea_TSpec_Fun."

Function LyTSpeci(S As TSpec, Specit$) As String() ' Return Ly of fst TSpeci of Spect.  Thw if there not such Spect
Dim I() As TSpeci: I = S.Itms
Dim M As TSpeci
Dim J%: For J = 0 To UbTSpeci(I)
    M = I(J)
'    If M.Specit = Specit Then LyILny (M.IxLny)
    
'    End If
    
Next
End Function
