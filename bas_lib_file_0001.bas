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
' Requiere referencia a Microsoft Office 11.0 Object Library (o superior 14.0 etc)
' -----------------------------------------------------------------------
' Funciones
' -----------------------------------------------------------------------
' file_0001_fNotExistingFileName
' file_0001_fDeleteFile
' file_0001_fCopyFile
' file_0001_fFileExists
' file_0001_fOpenDialogSelectFiles
' file_0001_fOpenDialogSaveAsFile
' 
' -----------------------------------------------------------------------
' Const msoFileDialogOpen = 1
' Const msoFileDialogSaveAs = 2
' Const msoFileDialogFilePicker = 3
' Const msoFileDialogFolderPicker = 4
' -----------------------------------------------------------------------
Function file_0001_fNotExistingFileName(txt As String)
  Dim ret As String
  ret = txt
  file_0001_fNotExistingFileName = ret
End Function
Sub file_0001_fDeleteFile(filename As String)
  If (file_0001_fFileExists(filename) = True) Then
    Kill filename
  End If
End Sub
Sub file_0001_fCopyFile(filename1 As String, filename2 As String)
  If (file_0001_fFileExists(filename1) = True) Then
    If (file_0001_fFileExists(filename2) = False) Then
      FileCopy filename1, filename2
    End If
  End If
End Sub
Function file_0001_fFileExists(filename As String)
  If Dir(filename) <> "" Then
    file_0001_fFileExists = True
  Else
    file_0001_fFileExists = False
  End If
End Function
Function file_0001_fOpenDialogSelectFiles(allow_multiselect As Boolean, title As String, default_file As String, AR_file_filter_name() As String, AR_file_filter_expr() As String, AR_file_selected() As String) As Boolean
  ' ---------------------------------------------------------------------
  ' Requires reference to Microsoft Office 11.0 Object Library (or greater 14.0 etc)
  ' ---------------------------------------------------------------------
  ' Ejemplo de llamada
  ' ---------------------------------------------------------------------
  ' ReDim AR_file_filter_name(1 To 1) As String
  ' ReDim AR_file_filter_expr(1 To 1) As String
  ' ReDim AR_file_selected(1 To 1) As String
  ' AR_file_filter_name(1) = "Excel files"
  ' AR_file_filter_expr(1) = "*.xls"
  ' ret = file_0001_fOpenDialogSelectFiles(True, "Please select one or more files", GLO_path, AR_file_filter_name(), AR_file_filter_expr(), AR_file_selected())
  ' ---------------------------------------------------------------------
  Dim i As Integer
  Dim ret As Boolean
  If GLO_filedialog_managed_by = "Office" Then
    'Dim fDialog As Office.FileDialog
  Else
    Dim fDialog As Object
    Set fDialog = Application.FileDialog(3)
  End If
  Dim varFile As Variant
  If GLO_filedialog_managed_by = "Office" Then
    'Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
  Else
    Set fDialog = Application.FileDialog(3)
  End If
  fDialog.title = title
  fDialog.AllowMultiSelect = allow_multiselect
  fDialog.Filters.Clear
  For i = 1 To UBound(AR_file_filter_name())
    fDialog.Filters.Add AR_file_filter_name(i), AR_file_filter_expr(i)
  Next i
  fDialog.InitialFileName = default_file
  ReDim AR_file_selected(1 To 1) As String
  If fDialog.Show = True Then
    i = 0
    For Each varFile In fDialog.SelectedItems
      i = i + 1
      ReDim Preserve AR_file_selected(1 To i) As String
      AR_file_selected(i) = varFile
    Next
    ret = True
  Else
    ret = False
  End If
  file_0001_fOpenDialogSelectFiles = ret
End Function
Function file_0001_fOpenDialogSaveAsFile(title As String, default_file As String, file_selected As String) As Boolean
  ' ---------------------------------------------------------------------
  ' Ejemplo de llamada
  ' ---------------------------------------------------------------------
  ' Dim file_selected As String
  ' ret = file_0001_fOpenDialogSaveAsFile("Please select output file", GLO_path, file_selected)
  ' ---------------------------------------------------------------------
  Dim ret As Boolean
  If GLO_filedialog_managed_by = "Office" Then
    'Dim fDialog As Office.FileDialog
  Else
    Dim fDialog As Object
    Set fDialog = Application.FileDialog(3)
  End If
  Dim varFile As Variant
  If GLO_filedialog_managed_by = "Office" Then
    'Set fDialog = Application.FileDialog(msoFileDialogSaveAs)
  Else
    Set fDialog = Application.FileDialog(2)
  End If
  fDialog.title = title
  fDialog.InitialFileName = default_file
  If fDialog.Show = True Then
    file_selected = fDialog.SelectedItems(1)
    ret = True
  Else
    ret = False
  End If
  file_0001_fOpenDialogSaveAsFile = ret
End Function

