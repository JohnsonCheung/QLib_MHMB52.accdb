Attribute VB_Name = "MxVb_X"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_X."
Private A$()
Sub ClrXX():                    Erase A:                End Sub
Sub XBox(S$):                   X Box(S):               End Sub
Sub XEnd():                     PushI A, "End":         End Sub
Sub XDrs(Drs As Drs):           PushIAy A, FmtDrs(Drs): End Sub
Sub XLn(Optional L$):           PushI A, L:             End Sub
Function XXLines$():  XXLines = JnCrLf(XX):             End Function
Function XX() As String()
XX = A
Erase A
End Function


Sub XTab(V)
If IsArray(V) Then
    X AmTab(V)
Else
    X vbTab & V
End If
End Sub

Sub X(V)
If IsArray(V) Then
    PushIAy A, V
Else
    PushI A, V
End If
End Sub
