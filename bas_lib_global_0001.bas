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
' Common global objects
' -----------------------------------------------------------------------
Global GLO_path As String
Global GLO_deploy_mode As Boolean
Global GLO_test_mode As Boolean
Global GLO_time_start As Variant
Global GLO_time_end As Variant
Global GLO_this_app_name As String
Global Const CTE_GLO_office_driver_app_excel As String = "EXCEL"
Global Const CTE_GLO_office_driver_app_access As String = "ACCESS"
Global Const CTE_GLO_office_driver_app_ppt As String = "PPT"
' -----------------------------------------------------------------------
' General global objects
' -----------------------------------------------------------------------
