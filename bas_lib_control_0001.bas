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
' control_0001_fStartBigProcess
' control_0001_fEndBigProcess
' 
' -----------------------------------------------------------------------
Sub control_0001_fStartBigProcess()
  Select Case GLO_office_driver_app
  Case CTE_GLO_office_driver_app_excel
  Case CTE_GLO_office_driver_app_access
  Case CTE_GLO_office_driver_app_ppt
  Case Else
  End Select
  'Screen.MousePointer = vbHourglass
  'Application.Cursor = xlWait
  db_0001_fFreeDBNoSelectQuery ("DELETE FROM t_log")
End Sub
Sub control_0001_fEndBigProcess()
  Select Case GLO_office_driver_app
  Case CTE_GLO_office_driver_app_excel
  Case CTE_GLO_office_driver_app_access
  Case CTE_GLO_office_driver_app_ppt
  Case Else
  End Select
  'Screen.MousePointer = vbDefault
  'Application.Cursor = xlDefault
End Sub
