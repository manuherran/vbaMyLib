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
' Requiere ejecutar Randomize al inicio del proyecto
' -----------------------------------------------------------------------
' Funciones
' -----------------------------------------------------------------------
' rand_0001_randDistContUnif1Ub
' rand_0001_randDistContUnif01
' rand_0001_randDistDiscUnifLbUb1
' rand_0001_randomNumbersString
' 
' 
' 
' 
' -----------------------------------------------------------------------
Function rand_0001_randDistContUnif1Ub(Ub As Double) As Double
  rand_0001_randDistContUnif1Ub = ((Ub - 1) * rand_0001_randDistContUnif01) + 1
End Function
Function rand_0001_randDistContUnif01() As Double
  rand_0001_randDistContUnif01 = Rnd
End Function
Function rand_0001_randDistDiscUnifLbUb1(Lb As Long, Ub As Long) As Long
'Devuleve un número al azar entero
'entre inicio y fin, inclusive ambos
'La probabilidad es igual para todos los números
'Se admiten valores negativos
  'Metodo 1
  'rand_0001_randDistDiscUnifLbUb1 = randDistDiscUnif1Ub1(Ub - Lb + 1) + Lb - 1
  'Metodo 2
  rand_0001_randDistDiscUnifLbUb1 = Int((Ub - Lb + 1) * rand_0001_randDistContUnif01 + Lb)
End Function
Function rand_0001_randomNumbersString(str_length As Integer) As String
  Dim ret As String
  ret = ""
  While Len(ret) < str_length
    ret = ret & CStr(rand_0001_randDistDiscUnifLbUb1(0, 9))
  Wend
  rand_0001_randomNumbersString = ret
End Function


