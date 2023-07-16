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
' ppt_0001_fCopySlideToTheEnd
' ppt_0001_fCreateLabel
' ppt_0001_fPrintArray2DAsPptTable
' ppt_0001_fCreateSlideWithExcelSheet
' ppt_0001_fConvertExcelWorkbookToPpt
' 
' -----------------------------------------------------------------------
Function ppt_0001_fCopySlideToTheEnd(template_page_index As Integer)
  Dim destiny_page_index As Integer
  Dim total_pages As Integer
  Dim prst1 As Presentation
  Dim sld1 As Slide
  ActivePresentation.Slides(template_page_index).Select
  ActivePresentation.Slides(template_page_index).Duplicate
  'La hoja duplicada se coloca a continuación de la original
  destiny_page_index = template_page_index + 1
  Set prst1 = ActivePresentation
  Set sld1 = prst1.Slides(destiny_page_index)
  total_pages = ActivePresentation.Slides.Count
  sld1.MoveTo total_pages
  destiny_page_index = total_pages
  ppt_0001_fCopySlideToTheEnd = destiny_page_index
End Function
Function ppt_0001_fCreateLabel(destiny_page_index As Integer, text As String, font_name As String, font_size As String, font_color As Long, font_bold As Boolean, left_px As Integer, top_px As Integer, width As Integer, height As Integer, text_alignment As String)
  Dim shape_index As Long
  ActivePresentation.Slides(destiny_page_index).Shapes.AddTextbox(msoTextOrientationHorizontal, left_px, top_px, width, height).Apply
  shape_index = ActivePresentation.Slides(destiny_page_index).Shapes.Count
  ActivePresentation.Slides(destiny_page_index).Shapes(shape_index).TextEffect.text = text
  ActivePresentation.Slides(destiny_page_index).Shapes(shape_index).TextFrame.TextRange.Font.Size = font_size
  ActivePresentation.Slides(destiny_page_index).Shapes(shape_index).TextFrame.TextRange.Font.Color = font_color
  ActivePresentation.Slides(destiny_page_index).Shapes(shape_index).TextFrame.TextRange.Font.Bold = font_bold 'msoFalse msoTrue
  If font_name <> "" Then
    ActivePresentation.Slides(destiny_page_index).Shapes(shape_index).TextFrame.TextRange.Font.Name = font_name
  End If
  If text_alignment = "center" Then
    ActivePresentation.Slides(destiny_page_index).Shapes(shape_index).TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter
  ElseIf text_alignment = "right" Then
    ActivePresentation.Slides(destiny_page_index).Shapes(shape_index).TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignRight
  Else
    ActivePresentation.Slides(destiny_page_index).Shapes(shape_index).TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft
  End If
End Function
Function ppt_0001_fPrintArray2DAsPptTable(page_index As Integer, AR_data() As Variant, p_left As Integer, p_top As Integer) As Integer
  Dim shape_index As Integer
  Dim rows As Integer
  Dim cols As Integer
  Dim i As Integer
  Dim j As Integer
  rows = UBound(AR_data, 1)
  cols = UBound(AR_data, 2)
  ActivePresentation.Slides(page_index).Select
  ActivePresentation.Slides(page_index).Shapes.AddTable NumRows:=rows, Numcolumns:=cols, left:=p_left, top:=p_top, width:=400, height:=40
  shape_index = ActivePresentation.Slides(page_index).Shapes.Count
  'Debug.Print ActivePresentation.Slides(page_index).Shapes.Count
  For j = 1 To cols
    ActivePresentation.Slides(page_index).Shapes(shape_index).Table.Columns(j).width = 50
  Next j
  For i = 1 To rows
    For j = 1 To cols
      If i = 1 Then
        ActivePresentation.Slides(page_index).Shapes(shape_index).Table.Cell(i, j).Shape.TextFrame.TextRange.Font.Size = 10
        ActivePresentation.Slides(page_index).Shapes(shape_index).Table.Cell(i, j).Shape.TextFrame.TextRange.Font.Bold = msoTrue
        ActivePresentation.Slides(page_index).Shapes(shape_index).Table.Cell(i, j).Shape.Fill.ForeColor.RGB = RGB(255, 0, 0)
        ActivePresentation.Slides(page_index).Shapes(shape_index).Table.Cell(i, j).Shape.TextFrame.TextRange.Font.Color = RGB(255, 255, 255)
      Else
        ActivePresentation.Slides(page_index).Shapes(shape_index).Table.Cell(i, j).Shape.TextFrame.TextRange.Font.Size = 10
        ActivePresentation.Slides(page_index).Shapes(shape_index).Table.Cell(i, j).Shape.TextFrame.TextRange.Font.Bold = msoFalse
        ActivePresentation.Slides(page_index).Shapes(shape_index).Table.Cell(i, j).Shape.Fill.ForeColor.RGB = RGB(255, 230, 230)
        ActivePresentation.Slides(page_index).Shapes(shape_index).Table.Cell(i, j).Shape.TextFrame.TextRange.Font.Color = RGB(0, 0, 0)
      End If
      
      ActivePresentation.Slides(page_index).Shapes(shape_index).Table.Cell(i, j).Shape.TextFrame.VerticalAnchor = msoAnchorMiddle
      ActivePresentation.Slides(page_index).Shapes(shape_index).Table.Cell(i, j).Shape.TextFrame.TextRange.text = AR_data(i, j)
      ActivePresentation.Slides(page_index).Shapes(shape_index).Table.Cell(i, j).Shape.TextFrame.TextRange.Font.Underline = msoFalse
      ActivePresentation.Slides(page_index).Shapes(shape_index).Table.Cell(i, j).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
      ActivePresentation.Slides(page_index).Shapes(shape_index).Table.Cell(i, j).Shape.Fill.Visible = msoTrue
      ActivePresentation.Slides(page_index).Shapes(shape_index).Table.Cell(i, j).Shape.Fill.Solid
      DoEvents
    Next j
  Next i
  ppt_0001_fPrintArray2DAsPptTable = shape_index
End Function
Sub ppt_0001_fCreateSlideWithExcelSheet(template_page_index As Integer, sheet_name As String)
  Dim destiny_page_index As Integer
  
  'Creamos una hoja nueva como copia de la primera
  destiny_page_index = ppt_0001_fCopySlideToTheEnd(template_page_index)
  
  'Añadimos textos en la hoja nueva
  ActivePresentation.Slides(destiny_page_index).Select
  
  'Añadimos el número de página en la hoja nueva
  ppt_0001_fCreateLabel destiny_page_index, "Página " & ActivePresentation.Slides.Count - 1, 12, RGB(0, 0, 0), False, 635, 510, 70, 20
  
  Dim firstRow As Long
  Dim firstCol As Long
  Dim lastRow As Long
  Dim lastCol As Long
  ReDim AR_data(1 To 1, 1 To 1) As Variant
  Dim i As Long
  Dim j As Long
  
  'Calculamos el rango con datos de esa hoja excel
  firstRow = 1
  firstCol = 1
  lastRow = 1
  lastCol = 1
  excel_0004_fCalculateRangeWithDataOfClosedExcelSheetCheckSmart GLO_excel_input_file, sheet_name, firstRow, firstCol, lastRow, lastCol
  
  'Cargamos la hoja excel en un array
  excel_0005_fLoadExcelSheetIntoArray GLO_excel_input_file, sheet_name, firstRow, firstCol, lastRow, lastCol, AR_data()
  'For i = 1 To UBound(AR_data, 1)
  '  For j = 1 To UBound(AR_data, 2)
  '    'Debug.Print "(" & i & "," & j & ")-->(" & AR_data(i, j) & ")"
  '    DoEvents
  '  Next j
  'Next i
  
  'Ignoramos las columnas de Plan y Eje (1 y 2) y también la última columna de comentarios
  Dim txt_plan As String
  Dim txt_eje As String
  txt_plan = "Plan: " & AR_data(2, 1)
  txt_eje = "Eje: " & AR_data(2, 2)
  
  'Añadimos el título en la hoja nueva
  ppt_0001_fCreateLabel destiny_page_index, txt_plan, 18, RGB(255, 0, 0), True, 10, 10, 400, 20
  ppt_0001_fCreateLabel destiny_page_index, txt_eje, 16, RGB(0, 0, 0), True, 10, 30, 400, 20
  
  'Realiza una copia del AR_data sobre el AR_data2, ambos de dimensión 2
  ReDim AR_data2(LBound(AR_data, 1) To UBound(AR_data, 1), LBound(AR_data, 2) To UBound(AR_data, 2) - 3) As Variant
  Dim L_n As Long
  Dim L_x As Long
  For L_n = LBound(AR_data, 1) To UBound(AR_data, 1)
    For L_x = 3 To UBound(AR_data, 2) - 1
      AR_data2(L_n, L_x - 2) = AR_data(L_n, L_x)
    Next L_x
  Next L_n
  
  'Pintamos el array en la slide como tabla
  ppt_0001_fPrintArray2DAsPptTable destiny_page_index, AR_data2()
  
  ActivePresentation.Save
End Sub
Sub ppt_0001_fConvertExcelWorkbookToPpt(filename As String)
  Dim template_page_index As Integer
  Dim i As Integer
  ReDim AR_sheet_names(1 To 1) As String
  template_page_index = 1
  excel_0005_fLoadExcelSheetNamesIntoArray filename, AR_sheet_names()
  For i = 1 To UBound(AR_sheet_names())
    ppt_0001_fCreateSlideWithExcelSheet template_page_index, AR_sheet_names(i)
    DoEvents
  Next i
End Sub

