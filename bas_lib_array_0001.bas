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
' array_0001_copyArray1DimIntoArray1Dim
' array_0001_copyArray2DimIntoArray2Dim
' array_0001_copyArray1DimIntoArray2Dim
' array_0001_array1DimIsEmpty 
' array_0001_fPrintDebugArray2D
' array_0001_fQueueInit
' array_0001_fQueuePush
' array_0001_fQueuePop
' array_0001_fQueueRemoveItem
' array_0001_array1DimIsEmpty
' 
' 
' -----------------------------------------------------------------------
Sub array_0001_copyArray1DimIntoArray1Dim(array1() As Variant, array2() As Variant)
  Dim i As Long
  ReDim array2(LBound(array1) To UBound(array1)) As String
  For i = LBound(array1, 1) To UBound(array1, 1)
    array2(i) = array1(i)
    DoEvents
  Next i
End Sub
Sub array_0001_copyArray2DimIntoArray2Dim(array1() As Variant, array2() As Variant)
  'Realiza una copia del Array1 sobre el Array2, ambos de dimensión 2
  Dim L_n As Long
  Dim L_x As Long
  ReDim array2(LBound(array1, 1) To UBound(array1, 1), LBound(array1, 2) To UBound(array1, 2)) As Variant
  For L_n = LBound(array1, 1) To UBound(array1, 1)
    For L_x = LBound(array1, 2) To UBound(array1, 2)
      array2(L_n, L_x) = array1(L_n, L_x)
    Next L_x
  Next L_n
End Sub
Sub array_0001_copyArray1DimIntoArray2Dim(array1() As Variant, array2() As Variant, ptr As Integer)
  Dim i As Long
  Dim lb_d1 As Long
  Dim ub_d1 As Long
  Dim lb_d2 As Long
  Dim ub_d2 As Long
  'No se pueden cambiar los límites de la primera dimension. Este bloque no tiene efecto
  If ptr < LBound(array2, 1) Then
    lb_d1 = ptr
  Else
    lb_d1 = LBound(array2, 1)
  End If
  If ptr > UBound(array2, 1) Then
    ub_d1 = ptr
  Else
    ub_d1 = UBound(array2, 1)
  End If
  'Este bloque si tiene efecto
  If LBound(array1, 1) < LBound(array2, 2) Then
    lb_d2 = LBound(array1, 1)
  Else
    lb_d2 = LBound(array2, 2)
  End If
  If UBound(array1, 1) > UBound(array2, 2) Then
    ub_d2 = UBound(array1, 1)
  Else
    ub_d2 = UBound(array2, 2)
  End If
  'No se pueden cambiar los límites de la primera dimension. La parte lb_d1 To ub_d1 debería mantenerse igual
  ReDim Preserve array2(lb_d1 To ub_d1, lb_d2 To ub_d2) As Variant
  For i = LBound(array1, 1) To UBound(array1, 1)
    array2(ptr, i) = array1(i)
    DoEvents
  Next i
End Sub
Sub array_0001_fPrintDebugArray2D(array1() As Variant)
  Dim L_n As Long
  Dim L_x As Long
  ReDim array2(LBound(array1, 1) To UBound(array1, 1), LBound(array1, 2) To UBound(array1, 2)) As Variant
  For L_n = LBound(array1, 1) To UBound(array1, 1)
    For L_x = LBound(array1, 2) To UBound(array1, 2)
      Debug.Print "(" & L_n & ", " & L_x & ") = " & array1(L_n, L_x)
    Next L_x
  Next L_n
End Sub
Function array_0001_fQueueInit(dato As Variant, myArray() As Variant)
  ReDim myArray(1 To 1) As Variant
  myArray(1) = dato
End Function
Function array_0001_fQueuePush(dato As Variant, myArray() As Variant)
  ReDim Preserve myArray(LBound(myArray) To UBound(myArray) + 1) As Variant
  myArray(UBound(myArray)) = dato
End Function
Function array_0001_fQueuePop(myArray() As Variant) As Variant
  If UBound(myArray) > LBound(myArray) Then
    lQueuePop = myArray(UBound(myArray))
    ReDim Preserve myArray(LBound(myArray) To UBound(myArray) - 1) As Long
  Else
    If UBound(myArray) = LBound(myArray) Then
      lQueuePop = myArray(UBound(myArray))
      ReDim myArray(1 To 1) As Long
    Else
      error_0001_fFatalError ("array_0001_fQueuePop")
    End If
  End If
End Function
Function array_0001_fQueueRemoveItem(indice As Long, myArray() As Variant) As Variant
  Dim i As Long
  Dim encontrado As Boolean
  encontrado = False
  For i = indice + 1 To UBound(myArray)
    myArray(i - 1) = myArray(i)
    encontrado = True
  Next i
  ReDim Preserve myArray(LBound(myArray) To UBound(myArray) - 1) As String
  If encontrado Then
    array_0001_fQueueRemoveItem = True
  Else
    array_0001_fQueueRemoveItem = False
  End If
End Function
Function array_0001_array1DimIsEmpty(myArray() As Variant) As Boolean
  Dim i As Long
  Dim ret As Boolean
  ret = True
  For i = LBound(myArray()) To UBound(myArray())
    If myArray(i) <> "" Then
      ret = False
      Exit For
    End If
  Next i
  array_0001_array1DimIsEmpty = ret
End Function
