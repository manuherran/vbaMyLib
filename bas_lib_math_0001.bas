' -----------------------------------------------------------------------
' vbaMyLib Version: 0.1.2 Release Date: 20170123
' © Copyright 2001-2023 Manu Herrán
' Free download source code:
' http://manuherran.com/
' -----------------------------------------------------------------------
Option Explicit
' -----------------------------------------------------------------------
' Tested with Access 2007
' -----------------------------------------------------------------------
' Funciones
' -----------------------------------------------------------------------
' math_0001_fScaleChange
' 
' 
' -----------------------------------------------------------------------
Function math_0001_fScaleChange(myValue, LbCurrentScale, UbCurrentScale, LbNewScale, UbNewScale) As Long
  ' Cambia de escala un número
  ' Por ejemplo, fScaleChange(0.9, 0.7, 1, 400, 700) = 600
  ' El numero 0.9 en una escala que va de 0.7 a 1 corresponde
  ' proporcionalmente con el valor solucion 600 en una escala que va de 400 a 700
  Dim ret As Long
  If UbCurrentScale = LbCurrentScale Then
    ret = UbNewScale
  Else
    ret = LbNewScale + ((myValue - LbCurrentScale) * ((UbNewScale - LbNewScale) / (UbCurrentScale - LbCurrentScale)))
  End If
  math_0001_fScaleChange = ret
End Function

