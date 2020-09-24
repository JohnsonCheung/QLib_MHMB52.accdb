Attribute VB_Name = "MxIde_Mth_CSub_zIntl_EnsCSubMsrc"
Option Compare Database
Option Explicit
Function MsrcyCSubPC() As Nly()
Dim M%: M = CPj.VBComponents.Count
Dim I%
Dim C As VBComponent: For Each C In CPj.VBComponents
    If I Mod 100 = 0 Then Debug.Print "MsrcyCSub: "; I; "of"; M; C.Name
    I = I + 1
    With SrcoptEnsCSub(C.CodeModule)
        If .Som Then
            PushNly MsrcyCSubPC, Nly(C.Name, .Ly)
        End If
    End With
Next
Debug.Print "NMsrcNew:"; NMsrcNew
End Function

