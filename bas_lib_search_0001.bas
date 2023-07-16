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
' search_0001_searchArrayItem
' 
' 
' -----------------------------------------------------------------------
Global Const CTE_MENOS_UNO_NO_ENCONTRADO = -1
Global Const CTE_BUSQUEDA_SECUENCIAL = 10
Global Const CTE_BUSQUEDA_BINARIA = 20
Function search_0001_searchArrayItem(elemento As Variant, myArray() As Variant, Lb As Long, Ub As Long, Optional matchCase, Optional searchAlgorithm) As Long
'-------------------------------------------------
'Tratamiento de parámetros opcionales. Por defecto:
'matchCase = CTE_MATCH_CASE
'searchAlgorithm = CTE_BUSQUEDA_SECUENCIAL
If IsMissing(matchCase) Then matchCase = CTE_MATCH_CASE
If IsMissing(searchAlgorithm) Then searchAlgorithm = CTE_BUSQUEDA_SECUENCIAL
'-------------------------------------------------
'Devuelve la posición
'Ejemplo de llamada
'casilla_a_borrar = searchArrayItem_s("Azul", miarr(), 1, 30)
'-------------------------------------------------
  Dim i As Long
  Select Case matchCase
  Case CTE_MATCH_CASE
      Select Case searchAlgorithm
      Case CTE_BUSQUEDA_SECUENCIAL
          'Búsqueda secuencial
          i = Lb
          While i <= Ub
              If Trim(myArray(i)) = Trim(elemento) Then
                  searchArrayItem_s = i
                  Exit Function
              End If
              i = i + 1
          Wend
          If i > Ub Then
              searchArrayItem_s = CTE_MENOS_UNO_NO_ENCONTRADO
          Else
              error_0001_fFatalError "Error en searchArrayItem_s"
          End If
      Case CTE_BUSQUEDA_BINARIA
          i = Lb + Int((Ub - Lb) / 2)
          If Trim(myArray(i)) = Trim(elemento) Then
              searchArrayItem_s = i
          Else
              If Lb = Ub Then
                  searchArrayItem_s = CTE_MENOS_UNO_NO_ENCONTRADO
              Else
                  If Lb + 1 = Ub Then
                      If Trim(myArray(Lb)) = Trim(elemento) Then
                          searchArrayItem_s = Lb
                      Else
                          If Trim(myArray(Ub)) = Trim(elemento) Then
                              searchArrayItem_s = Ub
                          Else
                              searchArrayItem_s = CTE_MENOS_UNO_NO_ENCONTRADO
                          End If
                      End If
                  Else
                      If Trim(myArray(i)) > Trim(elemento) Then
                          searchArrayItem_s = searchArrayItem_s(elemento, myArray(), Lb, i - 1, matchCase, searchAlgorithm)
                      Else
                          searchArrayItem_s = searchArrayItem_s(elemento, myArray(), i + 1, Ub, matchCase, searchAlgorithm)
                      End If
                  End If
              End If
          End If
      Case Else
          error_0001_fFatalError "Tipo de búsqueda no existente"
      End Select
  '===========================================================================
  Case CTE_NO_MATCH_CASE
  '===========================================================================
      Select Case searchAlgorithm
      Case CTE_BUSQUEDA_SECUENCIAL
          'Búsqueda secuencial
          i = Lb
          While i <= Ub
              If UCase(Trim(myArray(i))) = UCase(Trim(elemento)) Then
                  searchArrayItem_s = i
                  Exit Function
              End If
              i = i + 1
          Wend
          If i > Ub Then
              searchArrayItem_s = CTE_MENOS_UNO_NO_ENCONTRADO
          Else
              error_0001_fFatalError "Error en searchArrayItem_s"
          End If
      Case CTE_BUSQUEDA_BINARIA
          i = Lb + Int((Ub - Lb) / 2)
          If UCase(Trim(myArray(i))) = UCase(Trim(elemento)) Then
              searchArrayItem_s = i
          Else
              If Lb = Ub Then
                  searchArrayItem_s = CTE_MENOS_UNO_NO_ENCONTRADO
              Else
                  If Lb + 1 = Ub Then
                      If UCase(Trim(myArray(Lb))) = UCase(Trim(elemento)) Then
                          searchArrayItem_s = Lb
                      Else
                          If UCase(Trim(myArray(Ub))) = UCase(Trim(elemento)) Then
                              searchArrayItem_s = Ub
                          Else
                              searchArrayItem_s = CTE_MENOS_UNO_NO_ENCONTRADO
                          End If
                      End If
                  Else
                      If UCase(Trim(myArray(i))) > UCase(Trim(elemento)) Then
                          searchArrayItem_s = searchArrayItem_s(elemento, myArray(), Lb, i - 1, matchCase, searchAlgorithm)
                      Else
                          searchArrayItem_s = searchArrayItem_s(elemento, myArray(), i + 1, Ub, matchCase, searchAlgorithm)
                      End If
                  End If
              End If
          End If
      Case Else
          error_0001_fFatalError "Tipo de búsqueda no existente"
      End Select
  Case Else
      error_0001_fFatalError "Tipo de case no existente"
  End Select
End Function


