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
' ppt_shape_0001_fCreateShapeCircle
' ppt_shape_0001_fCreateShapeLabel
' ppt_shape_0001_fCreateCircleWithText
' ppt_shape_0001_fMoveShape
' 
' 
' 
' -----------------------------------------------------------------------
Function ppt_shape_0001_fCreateShapeCircle(page_index As Integer, fill_color As Long, line_color As Long, size As Integer, top_px As Integer, left_px As Integer)
  'On Error Resume Next
  'Ocasionalmente esta función se detiene en AddShape, sin embargo continua manualmente sin problemas. Vamos a probar con DoEvents
  Dim shape_index As Integer
  DoEvents
  ActiveWindow.Selection.SlideRange.Shapes.AddShape(msoShapeOval, top_px, left_px, size, size).Select
  DoEvents
  shape_index = ActivePresentation.Slides(page_index).Shapes.Count
  ActivePresentation.Slides(page_index).Shapes(shape_index).Fill.BackColor.RGB = fill_color
  ActivePresentation.Slides(page_index).Shapes(shape_index).Line.ForeColor.RGB = line_color
  ppt_shape_0001_fCreateShapeCircle = shape_index
End Function
Function ppt_shape_0001_fCreateShapeLabel(destiny_page_index As Integer, text As String, font_name As String, font_size As String, font_color As Long, font_bold As Boolean, left_px As Integer, top_px As Integer, width As Integer, height As Integer, text_alignment As String)
  Dim shape_index As Long
  ActivePresentation.Slides(destiny_page_index).Shapes.AddTextbox(msoTextOrientationHorizontal, left_px, top_px, width, height).Apply
  shape_index = ActivePresentation.Slides(destiny_page_index).Shapes.Count
  ActivePresentation.Slides(destiny_page_index).Shapes(shape_index).TextEffect.text = text
  ActivePresentation.Slides(destiny_page_index).Shapes(shape_index).TextFrame.TextRange.Font.size = font_size
  ActivePresentation.Slides(destiny_page_index).Shapes(shape_index).TextFrame.TextRange.Font.color = font_color
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
Function ppt_shape_0001_fCreateCircleWithText(page_index As Integer, solid As Boolean, color As Integer, size As String, top_px As Integer, left_px As Integer)
  Dim shape_index As Long
  Dim width As Integer
  Dim height As Integer
  width = 10
  height = 10
  ActivePresentation.Slides(page_index).Shapes.AddTextbox(msoTextOrientationHorizontal, left_px, top_px, width, height).Apply
  shape_index = ActivePresentation.Slides(page_index).Shapes.Count
  If font_size = "" Then
    font_size = 11
  End If
  ActivePresentation.Slides(page_index).Shapes(shape_index).TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter
  ActivePresentation.Slides(page_index).Shapes(shape_index).TextFrame.TextRange.Font.color = color
  ActivePresentation.Slides(page_index).Shapes(shape_index).TextFrame.TextRange.Font.Name = "Wingdings 2"
  ActivePresentation.Slides(page_index).Shapes(shape_index).TextFrame.TextRange.Font.size = font_size
  ActivePresentation.Slides(page_index).Shapes(shape_index).TextFrame.TextRange.Font.Bold = msoTrue
  If solid = True Then
    ActivePresentation.Slides(page_index).Shapes(shape_index).TextFrame.TextRange.text = "˜" 'Sólido
  Else
    ActivePresentation.Slides(page_index).Shapes(shape_index).TextFrame.TextRange.text = "š" 'Circunferencia
  End If
  ppt_shape_0001_fCreateCircleWithText = shape_index
End Function
Function ppt_shape_0001_fMoveShape(page_index As Integer, shape_index As Long, direction As Integer)
  Dim shape_x As Integer
  Dim shape_y As Integer
  shape_y = ActivePresentation.Slides(page_index).Shapes(shape_index).Top
  shape_x = ActivePresentation.Slides(page_index).Shapes(shape_index).Left
  Select Case direction
  Case 1
    shape_y = shape_y - GLO_step
  Case 2
    shape_x = shape_x + GLO_step
  Case 3
    shape_y = shape_y - GLO_step
  Case 4
    shape_x = shape_x - GLO_step
  End Select
  If shape_y < GLO_min_row Then
    shape_y = GLO_max_row
  End If
  If shape_y > GLO_max_row Then
    shape_y = GLO_min_row
  End If
  If shape_x < GLO_min_col Then
    shape_x = GLO_max_col
  End If
  If shape_x > GLO_max_col Then
    shape_x = GLO_min_col
  End If
  ActivePresentation.Slides(page_index).Shapes(shape_index).Top = shape_y
  ActivePresentation.Slides(page_index).Shapes(shape_index).Left = shape_x
End Function
