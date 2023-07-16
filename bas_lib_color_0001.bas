' -----------------------------------------------------------------------
' vbaMyLib Version: 0.1.2 Release Date: 20170123
' © Copyright 2001-2023 Manu Herrán
' Free download source code:
' http://manuherran.com/
' -----------------------------------------------------------------------
Option Explicit
' -----------------------------------------------------------------------
' Tested with Access 2003
' -----------------------------------------------------------------------
' Funciones
' -----------------------------------------------------------------------
' color_0001_fRandomColor
' 
' 
' -----------------------------------------------------------------------
Global Const CTE_schemeColorWhite As Integer = 1
Global Const CTE_schemeColorRed As Integer = 2
Global Const CTE_schemeColorGreen As Integer = 3
Global Const CTE_schemeColorBlue As Integer = 4
Global Const CTE_schemeColorYellow As Integer = 5
Global Const CTE_schemeColorPink As Integer = 6
Global Const CTE_schemeColorLightBlue As Integer = 7
Global Const CTE_schemeColorBlack As Integer = 8
Global Const CTE_schemeColorDarkRed As Integer = 16
Global Const CTE_schemeColorDarkGreen As Integer = 17
Global Const CTE_schemeColorDarkBlue As Integer = 18
Global Const CTE_schemeColorDarkBrown As Integer = 19
Global Const CTE_schemeColorLavander As Integer = 20
Global Const CTE_schemeColorSalmon As Integer = 29
Global Const CTE_schemeColorGray25 As Integer = 22
Global Const CTE_schemeColorGray40 As Integer = 55
Global Const CTE_schemeColorGray50 As Integer = 24
Global Const CTE_schemeColorGray80 As Integer = 73
Global Const CTE_schemeColorPink2 As Integer = 33
Global Const CTE_schemeColorOrange As Integer = 47
' -----------------------------------------------------------------------
' vbBlack
' vbWhite
' vbBlue
' vbCyan
' vbGreen
' vbMagenta
' vbRed
' vbYellow
' -----------------------------------------------------------------------
Function color_0001_fRandomColor()
  color_0001_fRandomColor = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
End Function


