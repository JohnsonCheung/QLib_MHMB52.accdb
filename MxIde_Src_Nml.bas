Attribute VB_Name = "MxIde_Src_Nml"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Nml."
Private Function WRmvAtrLines(Src$()) As String()
Dim Fm%
    Dim J%: For J = 0 To UB(Src)
        If Not HasPfx(Src(J), "Attribute ") Then
            Fm = J
            GoTo X
        End If
    Next
X:
WRmvAtrLines = AwFm(Src, Fm)
End Function
Function RmvClassHdrLines(XSrc$()) As String(): RmvClassHdrLines = WRmvAttrLines(WRmv4ClassLines(XSrc)): End Function
Private Function WRmv4ClassLines(XSrc$()) As String()
If Si(XSrc) = 0 Then Exit Function
If XSrc(0) = "VERSION 1.0 CLASS" Then
    WRmv4ClassLines = AwFm(XSrc, 4)
Else
    WRmv4ClassLines = XSrc
End If
End Function
Private Function WRmvAttrLines(XSrc$()) As String(): WRmvAttrLines = AwFm(XSrc, WNonAttrBix(XSrc)): End Function
Private Function WNonAttrBix&(XSrc$())
Dim J%: For J = 0 To UB(XSrc)
    If Not HasPfx(XSrc(J), "Attribute") Then WNonAttrBix = J: Exit Function
Next
WNonAttrBix = Si(XSrc)
End Function
