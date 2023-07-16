' -----------------------------------------------------------------------
' vbaMyLib Version: 0.1.2 Release Date: 20170123
' © Copyright 2001-2023 Manu Herrán
' Free download source code:
' http://manuherran.com/
' -----------------------------------------------------------------------
Option Explicit
' -----------------------------------------------------------------------
' Tested with Power Point 2007
' -----------------------------------------------------------------------
' Funciones
' -----------------------------------------------------------------------
' excel_0005_fLoadExcelSheetIntoArray
' excel_0005_fLoadExcelSheetNamesIntoArray
' 
' -----------------------------------------------------------------------
Sub excel_0005_fLoadExcelSheetIntoArray(excelBookFilename As String, excelSheetName As String, firstRow As Long, firstCol As Long, lastRow As Long, lastCol As Long, AR_data() As Variant)
  ' ---------------------------------------------------------------------
  ' Ejemplo de llamada
  ' ---------------------------------------------------------------------
  ' ReDim AR_data(1 To 1, 1 To 1) As Variant
  ' Dim firstRow As Long
  ' Dim firstCol As Long
  ' Dim lastRow As Long
  ' Dim lastCol As Long
  ' firstRow = 2
  ' firstCol = 2
  ' lastRow = 3
  ' lastCol = 6
  ' excel_0005_fLoadExcelSheetIntoArray "1.xls", "Hoja1", firstRow, firstCol, lastRow, lastCol
  ' Dim i As Long
  ' Dim j As Long
  ' For i = 1 To UBound(AR_data, 1)
  '   For j = 1 To UBound(AR_data, 2)
  '     Debug.Print "(" & i & "," & j & ")-->(" & AR_data(i, j) & ")"
  '   Next j
  ' Next i
  ' ---------------------------------------------------------------------
  Dim Obj_Excel As Object
  Dim Obj_Libro As Object
  Dim Obj_Hoja As Object
  Dim i As Long
  Dim j As Long
  ReDim AR_data(1 To lastRow - firstRow + 1, 1 To lastCol - firstCol + 1) As Variant
  If GLO_deploy_mode = False Then
    GLO_path = vba_0001_CalculatePath()
  End If
  Set Obj_Excel = CreateObject("Excel.Application")
  Set Obj_Libro = Obj_Excel.Workbooks.Open(GLO_path & "\" & excelBookFilename)
  Obj_Excel.Worksheets(excelSheetName).Activate
  Set Obj_Hoja = Obj_Libro.ActiveSheet
  For i = firstRow To lastRow
    For j = firstCol To lastCol
      AR_data(i - firstRow + 1, j - firstCol + 1) = Obj_Hoja.Cells(i, j)
    Next j
  Next i
  Set Obj_Hoja = Nothing
  Obj_Excel.Workbooks.Close
  Set Obj_Libro = Nothing
  Set Obj_Excel = Nothing
End Sub



Sub excel_0005_fLoadExcelSheetNamesIntoArray(excelBookFilename As String, AR_data() As String)
  ' ---------------------------------------------------------------------
  ' Ejemplo de llamada
  ' ---------------------------------------------------------------------
  ' ReDim AR_data(1 To 1) As String
  ' excel_0005_fLoadExcelSheetNamesIntoArray "1.xls", AR_data()
  ' ---------------------------------------------------------------------
  Dim Obj_Excel As Object
  Dim Obj_Libro As Object
  Dim Obj_Hoja As Object
  Dim i As Integer
  If GLO_deploy_mode = False Then
    GLO_path = vba_0001_CalculatePath()
  End If
  Set Obj_Excel = CreateObject("Excel.Application")
  Set Obj_Libro = Obj_Excel.Workbooks.Open(GLO_path & "\" & excelBookFilename)
  ReDim AR_data(1 To Obj_Excel.Worksheets.Count) As String
  For i = 1 To Obj_Excel.Worksheets.Count
    AR_data(i) = Obj_Excel.Worksheets(i).Name
  Next i
  Set Obj_Hoja = Nothing
  Obj_Excel.Workbooks.Close
  Set Obj_Libro = Nothing
  Set Obj_Excel = Nothing
End Sub

