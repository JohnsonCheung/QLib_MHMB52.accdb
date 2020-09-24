Attribute VB_Name = "MxIde_Fea_MD5"
Option Compare Text
Option Explicit

Function MD5M$(M As CodeModule):  MD5M = MD5(SrclM(M)): End Function
Function MD5P$(P As VBProject):   MD5P = MD5(SrclP(P)): End Function
Function MD5PC$():               MD5PC = MD5P(CPj):     End Function
Function MD5MC$():               MD5MC = MD5M(CMd):     End Function
