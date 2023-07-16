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
' error_0001_fFatalError
' 
' 
' -----------------------------------------------------------------------
Sub error_0001_fFatalError(msg As String)
  Debug.Print "Error: " & msg
  MsgBox "Error: " & msg, vbCritical
  'End
End Sub
