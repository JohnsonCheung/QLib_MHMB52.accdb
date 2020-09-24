Attribute VB_Name = "MxVb_Run_RunFun"
Option Compare Text
Option Explicit

Function SySyfunn(Syfunn$) As String(): SySyfunn = Eval(Syfunn & "()"): End Function
Function SySyfunnIf(Syfunn$) As String()
On Error Resume Next
SySyfunnIf = SySyfunn(Syfunn)
End Function
