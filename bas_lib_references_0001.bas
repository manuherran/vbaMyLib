' -----------------------------------------------------------------------
' vbaMyLib Version: 0.1.2 Release Date: 20170123
' © Copyright 2001-2023 Manu Herrán
' Free download source code:
' http://manuherran.com/
' -----------------------------------------------------------------------
Option Compare Database
'========================================================================
' Consulta: qryTestRefs
' SELECT LEFT(var,1) FROM t_config;
'========================================================================
' Macro: AutoExec
' Acción: EjecutarCódigo
' Datos: CheckRefs()
'========================================================================
' -----------------------------------------------------------------------
' Funciones
' -----------------------------------------------------------------------
' CheckRefs
' FixUpRefs
' 
' 
' -----------------------------------------------------------------------
Function CheckRefs()
  Dim db As Database, rs As Recordset
  Dim x
  Set db = CurrentDb
  On Error Resume Next
  ' Run the query qryTestRefs you created and trap for an error.
  Set rs = db.OpenRecordset("qryTestRefs", dbOpenDynaset)
  ' The if statement below checks for error 3075. If it encounters the
  ' error, it informs the user that it needs to fix the application.
  ' Error 3075 is the following:
  ' "Function isn't available in expressions in query expression..."
  ' Note: This function only checks for the error 3075. If you want it to
  ' check for other errors, you can modify the If statement. To have
  ' it check for any error, you can change it to the following:
  ' If Err.Number <> 0
   If Err.Number = 3075 Then
     MsgBox "This application has detected newer versions " _
            & "of required files on your computer. " _
            & "It may take several minutes to recompile " _
            & "this application."
     Err.Clear
     FixUpRefs
    End If
End Function
Sub FixUpRefs()
  Dim loRef As Access.Reference
  Dim intCount As Integer
  Dim intX As Integer
  Dim blnBroke As Boolean
  Dim strPath As String
  On Error Resume Next
  'Count the number of references in the database
  intCount = Access.References.Count
  'Loop through each reference in the database
  'and determine if the reference is broken.
  'If it is broken, remove the Reference and add it back.
  For intX = intCount To 1 Step -1
    Set loRef = Access.References(intX)
    With loRef
      blnBroke = .IsBroken
      If blnBroke = True Or Err <> 0 Then
        strPath = .FullPath
        With Access.References
          .Remove loRef
          .AddFromFile strPath
        End With
      End If
     End With
  Next
  Set loRef = Nothing
  ' Call a hidden SysCmd to automatically compile/save all modules.
  Call SysCmd(504, 16483)
End Sub

