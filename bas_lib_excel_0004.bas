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
' excel_0004_fCalculateRangeWithDataOfClosedExcelSheetCheckSmart
' excel_0004_fIsRangeWithDataOfOppenedSheetCheckAll
' 
' -----------------------------------------------------------------------
Sub excel_0004_fCalculateRangeWithDataOfClosedExcelSheetCheckSmart(excelBookFilename As String, excelSheetName As String, ByRef firstRowWithData As Long, ByRef firstColWithData As Long, ByRef lastRowWithData As Long, ByRef lastColWithData As Long)
  ' ---------------------------------------------------------------------
  ' Ejemplo de llamada
  ' ---------------------------------------------------------------------
  ' Dim firstRowWithData As Long
  ' Dim firstColWithData As Long
  ' Dim lastRowWithData As Long
  ' Dim lastColWithData As Long
  ' firstRowWithData = 1
  ' firstColWithData = 1
  ' lastRowWithData = 1
  ' lastColWithData = 1
  ' excel_0004_fCalculateRangeWithDataOfClosedExcelSheetCheckSmart "1.xls", "Hoja1", firstRowWithData, firstColWithData, lastRowWithData, lastColWithData
  ' Debug.Print "(" & firstRowWithData & "," & firstColWithData & ")-->(" & lastRowWithData & "," & lastColWithData & ")"
  ' ---------------------------------------------------------------------
  'firstRowWithData,firstColWithData son parámetros de entrada/salida
  'lastRowWithData,lastColWithData son parámetros de entrada/salida
  ' ---------------------------------------------------------------------
  Dim offset As Integer
  Dim max_offset As Integer
  Dim Obj_Excel As Object
  Dim Obj_Libro As Object
  Dim Obj_Hoja As Object
  If GLO_deploy_mode = False Then
    GLO_path = vba_0001_fCalculatePath()
  End If
  Set Obj_Excel = CreateObject("Excel.Application")
  Set Obj_Libro = Obj_Excel.Workbooks.Open(GLO_path & "\" & excelBookFilename)
  Obj_Excel.Worksheets(excelSheetName).Activate
  Set Obj_Hoja = Obj_Libro.ActiveSheet
  While Obj_Hoja.Cells(firstRowWithData, firstColWithData) = ""
    firstColWithData = firstColWithData + 1
    firstRowWithData = firstRowWithData + 1
    DoEvents
  Wend
  If firstColWithData > 1 Then
    While Obj_Hoja.Cells(firstRowWithData, firstColWithData - 1) <> ""
      firstColWithData = firstColWithData - 1
      DoEvents
    Wend
  End If
  If firstRowWithData > 1 Then
    While Obj_Hoja.Cells(firstRowWithData, firstColWithData - 1) <> ""
      firstRowWithData = firstRowWithData - 1
      DoEvents
    Wend
  End If
  lastRowWithData = firstRowWithData
  lastColWithData = firstColWithData
  max_offset = 3
  For offset = max_offset To 1 Step -1
    While Obj_Hoja.Cells(lastRowWithData + offset, lastColWithData + offset) <> ""
      lastColWithData = lastColWithData + offset
      lastRowWithData = lastRowWithData + offset
      DoEvents
    Wend
    While Obj_Hoja.Cells(lastRowWithData, lastColWithData + offset) <> ""
      lastColWithData = lastColWithData + offset
      DoEvents
    Wend
    While Obj_Hoja.Cells(lastRowWithData + offset, lastColWithData) <> ""
      lastRowWithData = lastRowWithData + offset
      DoEvents
    Wend
  Next offset
  'Ajustes
  If firstRowWithData > 1 Then
    While (excel_0004_fIsRangeWithDataOfOppenedSheetCheckAll(Obj_Hoja, 1, 1, firstRowWithData - 1, firstColWithData) = True)
      firstRowWithData = firstRowWithData - 1
      DoEvents
    Wend
  End If
  If firstColWithData > 1 Then
    While (excel_0004_fIsRangeWithDataOfOppenedSheetCheckAll(Obj_Hoja, 1, 1, firstRowWithData, firstColWithData - 1) = True)
      firstColWithData = firstColWithData - 1
      DoEvents
    Wend
  End If
  While (excel_0004_fIsRangeWithDataOfOppenedSheetCheckAll(Obj_Hoja, lastRowWithData + 1, 1, lastRowWithData + 1, lastColWithData) = True)
    lastRowWithData = lastRowWithData + 1
    DoEvents
  Wend
  While (excel_0004_fIsRangeWithDataOfOppenedSheetCheckAll(Obj_Hoja, 1, lastColWithData + 1, lastRowWithData, lastColWithData + 1) = True)
    lastColWithData = lastColWithData + 1
    DoEvents
  Wend
  Set Obj_Hoja = Nothing
  Obj_Excel.Workbooks.Close
  Set Obj_Libro = Nothing
  Set Obj_Excel = Nothing
End Sub
Function excel_0004_fIsRangeWithDataOfOppenedSheetCheckAll(Obj_Hoja As Object, startRow As Long, startCol As Long, endRow As Long, endCol As Long) As Boolean
  Dim existe As Boolean
  Dim i As Long
  Dim j As Long
  existe = False
  For i = startRow To endRow
    For j = startCol To endCol
      If Obj_Hoja.Cells(i, j) <> "" Then
        existe = True
      End If
    Next j
  Next i
  excel_0004_fIsRangeWithDataOfOppenedSheetCheckAll = existe
End Function
