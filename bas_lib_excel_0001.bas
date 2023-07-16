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
' excel_0001_fLoadExcelSheetIntoAccessDbTableSpecificCols
' excel_0001_fAddOneColAtTheEndOfTheSheet
' 
' 
' -----------------------------------------------------------------------
Function excel_0001_fLoadExcelSheetIntoAccessDbTableSpecificCols(excelWorkbookFilename As String, excelSheetName As String, dbTableName As String, excelStartReadingRow As Integer, deleteTableFirst As Boolean, totalNumberOfDbTableFieldsToLoad As Integer, ARR_excelColNumbersToLoad() As Integer, ARR_dbTableFieldNamesToLoad() As Variant)
  ' ---------------------------------------------------------------------
  ' Ejemplo de llamada
  ' ---------------------------------------------------------------------
  ' Dim ret As Boolean
  ' Dim totalNumberOfDbTableFieldsToLoad As Integer
  ' Dim excelStartReadingRow As Integer
  ' totalNumberOfDbTableFieldsToLoad = 2
  ' ReDim ARR_excelColNumbersToLoad(1 To totalNumberOfDbTableFieldsToLoad) As Integer
  ' ReDim ARR_dbTableFieldNamesToLoad(1 To totalNumberOfDbTableFieldsToLoad) As Variant
  ' No es necesario poner todos los campos de la tabla, pero los campos que se indiquen se han de cargar de alguna columna del excel. Si se quiere poner todos los campos aunque algunos no se vayan a cargar realmente, es necesario indicar alguna columna del excel que sea vacía como columna de carga de ese campo
  ' ARR_excelColNumbersToLoad(1) = 1
  ' ARR_excelColNumbersToLoad(2) = 2
  ' ARR_dbTableFieldNamesToLoad(1) = "id"
  ' ARR_dbTableFieldNamesToLoad(2) = "name"
  ' excelStartReadingRow = 2
  ' ret = excel_0001_fLoadExcelSheetIntoAccessDbTableSpecificCols(GLO_path & "\" & "Libro1.xls", "Hoja1", "t_tabla1", excelStartReadingRow, True, totalNumberOfDbTableFieldsToLoad, ARR_excelColNumbersToLoad(), ARR_dbTableFieldNamesToLoad())
  ' ---------------------------------------------------------------------
  Dim ret As Boolean
  Dim Obj_Excel As Object
  Dim Obj_Libro As Object
  Dim Obj_Hoja As Object
  Dim cont_f As Long
  Dim cont_c As Long
  Dim txt_debug As String
  Dim asignar_valor As Boolean
  If deleteTableFirst Then
    db_0001_fFreeDBNoSelectQuery ("DELETE FROM " & dbTableName & ";")
  End If
  Dim rst As Recordset
  Set rst = CurrentDb.OpenRecordset("SELECT * FROM " & dbTableName & ";", dbOpenDynaset)
  ReDim ARR_fila(1 To 1) As Variant
  ReDim ARR_fila(1 To totalNumberOfDbTableFieldsToLoad) As Variant
  Set Obj_Excel = CreateObject("Excel.Application")
  Obj_Excel.Workbooks.Open (excelWorkbookFilename)
  Obj_Excel.Worksheets(excelSheetName).Activate
  If Val(Obj_Excel.Application.Version) >= 8 Then
    Set Obj_Hoja = Obj_Excel.ActiveSheet
  Else
    Set Obj_Hoja = Obj_Excel
  End If
  cont_f = 1
  'Leo cabecera
  'txt_debug = ""
  For cont_c = 1 To totalNumberOfDbTableFieldsToLoad
    ARR_fila(cont_c) = Obj_Hoja.Cells(cont_f, ARR_excelColNumbersToLoad(cont_c))
    ARR_fila(cont_c) = Trim(ARR_fila(cont_c))
    'txt_debug = txt_debug & ARR_fila(cont_c) & " "
  Next cont_c
  'Debug.Print txt_debug
  'Proceso registros
  cont_f = excelStartReadingRow - 1
  'While ARR_fila(1) <> ""
  While Not array_0001_array1DimIsEmpty(ARR_fila())
    cont_f = cont_f + 1
    'txt_debug = ""
    rst.AddNew
    For cont_c = 1 To totalNumberOfDbTableFieldsToLoad
      asignar_valor = True
      ARR_fila(cont_c) = Obj_Hoja.Cells(cont_f, ARR_excelColNumbersToLoad(cont_c))
      ARR_fila(cont_c) = Trim(ARR_fila(cont_c))
      'Omito los datos que no llegan en el formato esperado, por ejemplo
      'Caso 1: Si espero un campo de tipo integer, y no llega un numérico. Esto ocurre cuando es un registro vacío en mitad de fichero, pero también puede ocurrir en cualquier punto
      If rst.Fields(ARR_dbTableFieldNamesToLoad(cont_c)).Type = CTE_VB_DataType_Integer Then
        If Not IsNumeric(ARR_fila(cont_c)) Then
          asignar_valor = False
        End If
      'Caso 2: Si espero un campo de tipo entero integer, y no llega un numérico. Esto ocurre cuando es un registro vacío en mitad de fichero, pero también puede ocurrir en cualquier punto
      ElseIf rst.Fields(ARR_dbTableFieldNamesToLoad(cont_c)).Type = CTE_VB_DataType_Double Then
        If Not IsNumeric(ARR_fila(cont_c)) Then
          asignar_valor = False
        End If
      'Caso 3: Si espero un campo de tipo fecha, y no llega una fecha, por ejemplo, si llega cadena vacía. Esto ocurre siempre al final de la hoja excel
      ElseIf rst.Fields(ARR_dbTableFieldNamesToLoad(cont_c)).Type = CTE_VB_DataType_Date Then
        If Not IsDate(ARR_fila(cont_c)) Then
          asignar_valor = False
        End If
      End If
      'Otros casos: ...
      If asignar_valor = True Then
        rst.Fields(ARR_dbTableFieldNamesToLoad(cont_c)).Value = ARR_fila(cont_c)
      End If
      'txt_debug = txt_debug & ARR_fila(cont_c) & " "
    Next cont_c
    'If ARR_fila(1) <> "" Then
    If Not array_0001_array1DimIsEmpty(ARR_fila()) Then
      rst.Update
      'Debug.Print cont_f & " - " & txt_debug
    End If
    DoEvents
  Wend
  Set Obj_Hoja = Nothing
  Obj_Excel.Workbooks.Close
  Set Obj_Libro = Nothing
  Set Obj_Excel = Nothing
  ret = True
  excel_0001_fLoadExcelSheetIntoAccessDbTableSpecificCols = ret
End Function
Function excel_0001_fAddOneColAtTheEndOfTheSheet(excelWorkbookFilename As String, excelSheetName As String, firstRow As String, nextRows As String)
  Dim firstRowWithData As Long
  Dim firstColWithData As Long
  Dim lastRowWithData As Long
  Dim lastColWithData As Long
  Dim i As Long
  firstRowWithData = 1
  firstColWithData = 1
  lastRowWithData = 1
  lastColWithData = 1
  excel_0004_fCalculateRangeWithDataOfClosedExcelSheetCheckSmart excelWorkbookFilename, excelSheetName, firstRowWithData, firstColWithData, lastRowWithData, lastColWithData
  Dim Obj_Excel As Object
  Dim Obj_Libro As Object
  Dim Obj_Hoja As Object
  If GLO_deploy_mode = False Then
    GLO_path = vba_0001_fCalculatePath()
  End If
  Set Obj_Excel = CreateObject("Excel.Application")
  Set Obj_Libro = Obj_Excel.Workbooks.Open(GLO_path & "\" & excelWorkbookFilename)
  Obj_Excel.Worksheets(excelSheetName).Activate
  Set Obj_Hoja = Obj_Libro.ActiveSheet
  Obj_Hoja.Cells(1, lastColWithData + 1) = firstRow
  For i = 2 To lastRowWithData
    Obj_Hoja.Cells(i, lastColWithData + 1) = nextRows
  Next i
  Obj_Libro.Save
  Obj_Excel.Quit
  Set Obj_Hoja = Nothing
  Obj_Excel.Workbooks.Close
  Set Obj_Libro = Nothing
  Set Obj_Excel = Nothing
End Function
