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
' excel_0002_fAllDataOfTwoExcelSheetsAreEqual
' excel_0002_fCompareTwoExcelSheetsAndChangeColorOfNotEqualCells
' excel_0002_fCreateStringsTableAndLoadAllExcelSheetIntoAccessDbTable
' excel_0002_fLoadExcelSheetIntoExistingAccessDbTable
' 
' -----------------------------------------------------------------------
Function excel_0002_fAllDataOfTwoExcelSheetsAreEqual(excelWorkbookFilename1 As String, excelSheetName1 As String, excelWorkbookFilename2 As String, excelSheetName2 As String, max_cols_to_check as Integer) As Boolean
  ' ---------------------------------------------------------------------
  ' Ejemplo de llamada
  ' ---------------------------------------------------------------------
  ' Dim ret As Boolean
  ' ret = excel_0002_fAllDataOfTwoExcelSheetsAreEqual("1.xls", "Hoja1", "2.xls", "Hoja1", 30)
  ' Debug.Print ret
  ' ---------------------------------------------------------------------
  Dim ret As Boolean
  Dim iguales As Boolean
  Dim sql As String
  Dim t1 As String
  Dim t2 As String
  t1 = "_OLD_t_temp_101_DELETE_ME"
  t2 = "_OLD_t_temp_102_DELETE_ME"
  sql = db_0001_fBuildSqlForCreateBigTableMemo(t1, max_cols_to_check)
  db_0001_fFreeDBNoSelectQuery (sql)
  sql = db_0001_fBuildSqlForCreateBigTableMemo(t2, max_cols_to_check)
  db_0001_fFreeDBNoSelectQuery (sql)
  ret = excel_0002_fCreateStringsTableAndLoadAllExcelSheetIntoAccessDbTable(excelWorkbookFilename1, excelSheetName1, t1, max_cols_to_check)
  ret = excel_0002_fCreateStringsTableAndLoadAllExcelSheetIntoAccessDbTable(excelWorkbookFilename2, excelSheetName2, t2, max_cols_to_check)
  iguales = db_0001_fAllDataOfTwoDbTablesAreEqual(t1, t2)
  excel_0001_fAllDataOfTwoExcelSheetsAreEqual = iguales
End Function
Function excel_0002_fCompareTwoExcelSheetsAndChangeColorOfNotEqualCells(excelWorkbookFilename1 As String, excelSheetName1 As String, excelWorkbookFilename2 As String, excelSheetName2 As String) As Boolean
  ' ---------------------------------------------------------------------
  ' Ejemplo de llamada
  ' ---------------------------------------------------------------------
  ' Dim ret As Boolean
  ' ret = excel_0002_fCompareTwoExcelSheetsAndChangeColorOfNotEqualCells("1.xls", "hoja1", "2.xls", "hoja1")
  ' ---------------------------------------------------------------------
  Dim Obj_Excel As Object
  Dim Obj_Libro1 As Object
  Dim Obj_Libro2 As Object
  Dim Obj_Hoja1 As Object
  Dim Obj_Hoja2 As Object
  Dim cont_f As Integer
  Dim cont_c As Integer
  Dim max_row As Integer
  Dim max_col As Integer
  max_row = 1000
  max_col = 15
  ReDim ARR_fila1(1 To max_col) As Variant
  ReDim ARR_fila2(1 To max_col) As Variant
  Set Obj_Excel = CreateObject("Excel.Application")
  If GLO_deploy_mode = False Then
    GLO_path = Left(CurrentDb.Name, InStrRev(CurrentDb.Name, "\"))
  End If
  Set Obj_Libro1 = Obj_Excel.Workbooks.Open(GLO_path & "\" & excelWorkbookFilename1)
  Set Obj_Libro2 = Obj_Excel.Workbooks.Open(GLO_path & "\" & excelWorkbookFilename2)
  Obj_Libro1.Worksheets(excelSheetName1).Activate
  Obj_Libro2.Worksheets(excelSheetName2).Activate
  Set Obj_Hoja1 = Obj_Libro1.ActiveSheet
  Set Obj_Hoja2 = Obj_Libro2.ActiveSheet
  'Obj_Excel1.Worksheets(excelSheetName1).Activate
  'Set Obj_Hoja1 = Obj_Excel1.ActiveSheet
  'Obj_Excel2.Worksheets(excelSheetName2).Activate
  'Set Obj_Hoja2 = Obj_Excel2.ActiveSheet
  'Proceso registros
  For cont_f = 1 To max_row
    For cont_c = 1 To max_col
      ARR_fila1(cont_c) = Obj_Hoja1.Cells(cont_f, cont_c)
      ARR_fila2(cont_c) = Obj_Hoja2.Cells(cont_f, cont_c)
      If ARR_fila1(cont_c) <> ARR_fila2(cont_c) Then
        Obj_Hoja1.Cells(cont_f, cont_c).interior.color = vbYellow
        Obj_Hoja2.Cells(cont_f, cont_c).interior.color = vbYellow
      End If
      DoEvents
    Next cont_c
    Debug.Print cont_f
  Next cont_f
  Obj_Libro1.Save
  Obj_Libro2.Save
  Obj_Excel.Workbooks.Close
  Set Obj_Hoja1 = Nothing
  Set Obj_Hoja2 = Nothing
  Set Obj_Libro1 = Nothing
  Set Obj_Libro2 = Nothing
  Set Obj_Excel = Nothing
End Function
Function excel_0002_fCreateStringsTableAndLoadAllExcelSheetIntoAccessDbTable(excelWorkbookFilename As String, excelSheetName As String, dbTableName As String, totalNumberOfDbTableFieldsToLoad as Integer)
  Dim ret As Boolean
  Dim excelStartReadingRow As Integer
  Dim i As Integer
  'Carga toda la excel en una tabla generica de tipos string
  ReDim ARR_excelColNumbersToLoad(1 To totalNumberOfDbTableFieldsToLoad) As Integer
  ReDim ARR_dbTableFieldNamesToLoad(1 To totalNumberOfDbTableFieldsToLoad) As Variant
  For i = 1 To totalNumberOfDbTableFieldsToLoad
    ARR_excelColNumbersToLoad(i) = i
    ARR_dbTableFieldNamesToLoad(i) = "field_" & CStr(i)
  Next i
  excelStartReadingRow = 1
  ret = excel_0001_fLoadExcelSheetIntoAccessDbTableSpecificCols(excelWorkbookFilename, excelSheetName, dbTableName, excelStartReadingRow, True, totalNumberOfDbTableFieldsToLoad, ARR_excelColNumbersToLoad(), ARR_dbTableFieldNamesToLoad())
End Function
Function excel_0002_fLoadExcelSheetIntoExistingAccessDbTable(excelWorkbookFilename As String, excelSheetName As String, dbTableName As String, excelStartReadingRow As Integer, totalNumberOfDbTableFieldsToLoad As Integer, deleteTableFirst As Boolean)
  Dim ret As Boolean
  Dim i As Integer
  ReDim ARR_excelColNumbersToLoad(1 To totalNumberOfDbTableFieldsToLoad) As Integer
  ReDim ARR_dbTableFieldNamesToLoad(1 To totalNumberOfDbTableFieldsToLoad) As Variant
  For i = 1 To totalNumberOfDbTableFieldsToLoad
    ARR_excelColNumbersToLoad(i) = i
    ARR_dbTableFieldNamesToLoad(i) = i - 1
  Next i
  ret = excel_0001_fLoadExcelSheetIntoAccessDbTableSpecificCols(excelWorkbookFilename, excelSheetName, dbTableName, excelStartReadingRow, deleteTableFirst, totalNumberOfDbTableFieldsToLoad, ARR_excelColNumbersToLoad(), ARR_dbTableFieldNamesToLoad())
End Function
