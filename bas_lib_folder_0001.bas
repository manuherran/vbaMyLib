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
' folder_0001_fAllFileNamesOfFolderName
' folder_0001_fOpenDialogSelectFolder
' 
' -----------------------------------------------------------------------
Function folder_0001_fAllFileNamesOfFolderName(myPath As String, AR_files() As Variant)
  ' -------------------------------------------------------------------
  ' Ejemplo de llamada:
  ' -------------------------------------------------------------------
  ' folder_0001_fAllFileNamesOfFolderName myPath, AR_files()
  ' For Each myFileIn In AR_files
  '   Debug.Print myFileIn
  ' Next myFileIn
  ' -------------------------------------------------------------------
  Erase AR_files()
  Dim cont As Long
  Dim MyObj As Object, MySource As Object, myFile As Variant
  Set MyObj = CreateObject("Scripting.FileSystemObject")
  Set MySource = MyObj.GetFolder(myPath)
  cont = 0
  For Each myFile In MySource.Files
    cont = cont + 1
    If cont = 1 Then
      array_0001_fQueueInit myFile.Name, AR_files
    Else
      array_0001_fQueuePush myFile.Name, AR_files
    End If
  Next myFile
End Function
Function folder_0001_fOpenDialogSelectFolder(title As String, default_folder As String, folder_selected As String) As Boolean
  ' ---------------------------------------------------------------------
  ' Requires reference to Microsoft Office 11.0 Object Library (or greater 14.0 etc)
  ' ---------------------------------------------------------------------
  ' Ejemplo de llamada
  ' ---------------------------------------------------------------------
  ' Dim folder_selected As String
  ' ret = folder_0001_fOpenDialogSelectFolder("Please select folder", GLO_path, folder_selected)
  ' ---------------------------------------------------------------------
  Dim ret As Boolean
  Dim fDialog As Office.FileDialog
  Dim varFile As Variant
  Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
  fDialog.title = title
  fDialog.Filters.Clear
  fDialog.InitialFileName = default_file
  ReDim AR_file_selected(1 To 1) As String
  If fDialog.Show = True Then
    folder_selected = fDialog.SelectedItems(1)
    ret = True
  Else
    ret = False
  End If
  folder_0001_fOpenDialogSelectFolder = ret
End Function
