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
' path_0001_fString2FileName
' path_0001_pathExists
' 
' -----------------------------------------------------------------------
Function path_0001_fString2FileName(txt As String)
  Dim ret As String
  ret = txt
  ret = str_0001_fStringMultilineToOneLineTrim(ret)
  ret = Replace(ret, "/", " ")
  ret = Replace(ret, ":", " ")
  ret = Replace(ret, vbTab, " ")
  ret = Replace(ret, vbCr, " ")
  ret = Replace(ret, vbLf, " ")
  ret = Replace(ret, vbCrLf, " ")
  ret = str_0001_fStringMultilineToOneLineTrim(ret)
  ret = Replace(ret, " ", "_")
  path_0001_fString2FileName = ret
End Function
Function path_0001_pathExists(Path As String) As Boolean
On Error GoTo noexiste
  If (GetAttr(Path) And vbDirectory) = vbDirectory Then
    path_0001_pathExists = True
  Else
    path_0001_pathExists = False
  End If
  Exit Function
noexiste:
  path_0001_pathExists = False
End Function

