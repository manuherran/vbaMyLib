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
' excel_0003_fKillAllExcelInstances
' 
' 
' 
' -----------------------------------------------------------------------
Sub excel_0003_fKillAllExcelInstances()
  On Error Resume Next
  Dim strComputer
  Dim objWMIService
  Dim colProcesses
  Dim objProcess
  Dim Response
  strComputer = "."
  Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
  Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'EXCEL.EXE'")
  Do While (colProcesses.Count <> 0)
    For Each objProcess In colProcesses
      objProcess.Terminate
    Next
    Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'EXCEL.EXE'")
  Loop
End Sub
