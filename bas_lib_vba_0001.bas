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
' vba_0001_fInit
' vba_0001_fCalculatePath
' vba_0001_fExitOfficeApp
' vba_0001_fCloseFormAndExit
' vba_0001_fStopVba
'
' -----------------------------------------------------------------------
Global GLO_office_driver_app As String
Global GLO_office_driver_version As String
Global GLO_excel_output_file_extension As String
Global GLO_filedialog_managed_by As String
' -----------------------------------------------------------------------
Global Const CTE_GLO_office_driver_app_excel As String = "EXCEL"
Global Const CTE_GLO_office_driver_app_access As String = "ACCESS"
Global Const CTE_GLO_office_driver_app_ppt As String = "PPT"
' -----------------------------------------------------------------------
Global Const CTE_VB_DataType_Integer As Integer = 4
Global Const CTE_VB_DataType_Double As Integer = 7
Global Const CTE_VB_DataType_Date As Integer = 8
Global Const CTE_VB_DataType_String As Integer = 10
' -----------------------------------------------------------------------
Sub vba_0001_fInit()
  '----------------------------------------------------------------------
  ' OFFICE
  ' GLO_office_driver_app = CTE_GLO_office_driver_app_excel
  '----------------------------------------------------------------------
  'GLO_filedialog_managed_by = "Office"
  GLO_filedialog_managed_by = "Application"

  '----------------------------------------------------------------------
  ' EXCEL
  ' GLO_office_driver_app = CTE_GLO_office_driver_app_excel
  '----------------------------------------------------------------------
  
  '----------------------------------------------------------------------
  ' ACCESS
  GLO_office_driver_app = CTE_GLO_office_driver_app_access
  GLO_office_driver_version = Application.Version
  Application.SetOption "Auto compact", True
  '----------------------------------------------------------------------
  
  '----------------------------------------------------------------------
  ' POWERPOINT
  ' GLO_office_driver_app = CTE_GLO_office_driver_app_ppt
  '----------------------------------------------------------------------

  Select Case GLO_office_driver_version
  Case "11.0"
    GLO_excel_output_file_extension = ".xls"
  Case "14.0"
    GLO_excel_output_file_extension = ".xlsx"
  Case Else
    GLO_excel_output_file_extension = ".xls"
  End Select
End Sub
Function vba_0001_fCalculatePath()
  Dim ret As String
  If GLO_deploy_mode = False Then
    vba_0001_fInit
  End If
  Select Case GLO_office_driver_app
  Case CTE_GLO_office_driver_app_excel
    'ret = Application.ActiveWorkbook.Path
  Case CTE_GLO_office_driver_app_access
    ret = Left(CurrentDb.Name, InStrRev(CurrentDb.Name, "\"))
    If Right(ret, 1) = "\" Then
      ret = Left(ret, Len(ret) - 1)
    End If
  Case CTE_GLO_office_driver_app_ppt
    'ret = ActivePresentation.Path
  Case Else
  End Select
  vba_0001_fCalculatePath = ret
End Function
Sub vba_0001_fExitOfficeApp()
  DoCmd.Quit
End Sub
Sub vba_0001_fCloseFormAndExit()
  DoCmd.Close acForm, "Form_FRM_MENU_PRINCIPAL"
  End
End Sub
Sub vba_0001_fStopVba()
  End
End Sub
