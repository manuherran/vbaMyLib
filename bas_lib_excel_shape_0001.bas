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
' shape_0001_fPrintCircleRowPxColPx
' shape_0001_fPrintCircleRowCol
' shape_0001_fTestShapes
' shape_0001_fInsertShapeFlashExplosion
' 
' 
' -----------------------------------------------------------------------
Sub shape_0001_fPrintCircleRowPxColPx(rowPx As Integer, colPx As Integer, radio As Integer, color As Integer)
  Dim myDocument As Object
  Dim myShapes As Object
  Dim myShape As Object
  Set myDocument = ActiveSheet
  Set myShapes = myDocument.Shapes
  Set myShape = myShapes.AddShape(msoShapeOval, colPx, rowPx, radio, radio)
  myShape.Fill.ForeColor.SchemeColor = color
End Sub
Sub shape_0001_fPrintCircleRowCol(row As Integer, col As Integer, radio As Integer, color As Integer)
  Dim rowPx As Integer
  Dim colPx As Integer
  rowPx = row * Range(Cells(1, 1), Cells(row - 1, 1)).Height / 3
  colPx = col * Range(Cells(1, 1), Cells(1, col - 1)).Width / 3
  shape_0001_fPrintCircleRowPxColPx rowPx, colPx, radio, color
End Sub
Sub shape_0001_fTestShapes()
  Dim i As Integer
  Dim row As Integer
  Dim col As Integer
  'shape_0001_fPrintCircleRowPxColPx 300, 300, 10, CTE_schemeColorGray40
  For row = 2 To 6
    For col = 2 To 6
      shape_0001_fPrintCircleRowCol row, col, 10, CTE_schemeColorRed
    Next col
  Next row
  If 1 = 2 Then
    For i = 41 To 50
      Cells(i, 5).Value = i
      shape_0001_fPrintCircleRowPxColPx 400, i * 10, 10, i
    Next i
  End If
End Sub
Sub shape_0001_fInsertShapeFlashExplosion(text As String)
  ActiveSheet.Shapes.AddShape(msoShapeExplosion2, 407.25, 162#, 525#, 409.5).Select
  Selection.ShapeRange.Fill.ForeColor.SchemeColor = 13
  Selection.ShapeRange.Fill.Visible = msoTrue
  Selection.ShapeRange.Fill.Solid
  Selection.Characters.text = text
  With Selection
  .HorizontalAlignment = xlCenter
  .VerticalAlignment = xlCenter
  .ReadingOrder = xlContext
  .Orientation = xlHorizontal
  .AutoSize = False
  End With
  With Selection.Font
  .Name = "BMWTypeRegular"
  .FontStyle = "Bold"
  .Size = 20
  .Strikethrough = False
  .Superscript = False
  .Subscript = False
  .OutlineFont = False
  .Shadow = False
  .Underline = xlUnderlineStyleNone
  .ColorIndex = xlAutomatic
  End With
  Range("A1").Select
End Sub
