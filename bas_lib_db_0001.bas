' -----------------------------------------------------------------------
' vbaMyLib Version: 0.1.2 Release Date: 20170123
' © Copyright 2001-2023 Manu Herrán
' Free download source code:
' http://manuherran.com/
' -----------------------------------------------------------------------
Option Explicit
' -----------------------------------------------------------------------
' Tested with Access
' - Access 2003: Yes
' - Access 2007: Yes / No
' - Access 2010: Yes / No
' Tested with Excel
' - Excel 2003: Yes / No
' - Excel 2007: Yes / No
' - Excel 2010: Yes / No
' -----------------------------------------------------------------------
' Funciones
' -----------------------------------------------------------------------
' db_0001_fFreeDBNoSelectQuery
' db_0001_fSqlOneFieldQuery
' db_0001_fDBRowExists
' db_0001_fPrepareTxtFieldToBeInserted
' db_0001_fBuildSqlForCreateBigTableMemo
' db_0001_fAllDataOfTwoDbTablesAreEqual
' db_0001_fFreeDBNRowQuery
' db_0001_fFreeDBNRowNotArrayQuery
' db_0001_fSaveLog
' 
' 
' 
' 
' -----------------------------------------------------------------------
Global Const CTE_db_temp_query As String = "___DELETE_ME___BORRAME___"
Global Const CTE_db_no_record As String = "***NO-RECORD***"
Global Const CTE_db_output_format_4 As Integer = 4
Global Const CTE_db_output_format_14 As Integer = 14
Global Const CTE_db_output_format_15 As Integer = 15
Sub db_0001_fFreeDBNoSelectQuery(strSQL As String)
On Error GoTo Err_db_0001_fFreeDBNoSelectQuery
  Dim dbs As DAO.Database
  Dim qdf As DAO.QueryDef
  Set dbs = CurrentDb
  Set qdf = dbs.CreateQueryDef(CTE_db_temp_query, strSQL)
  qdf.Execute
  dbs.QueryDefs.Delete (CTE_db_temp_query)
  dbs.Close
Exit_db_0001_fFreeDBNoSelectQuery:
  Exit Sub
Err_db_0001_fFreeDBNoSelectQuery:
  Dim test As String
  test = ""
  dbs.Close
  MsgBox Err.Description, vbCritical
  MsgBox "El proceso va a finalizar. Se recomienda borrar manualmente la consulta " & CTE_db_temp_query & " e iniciar el proceso de nuevo.", vbInformation
  error_0001_fFatalError ""
  End
  'Resume Exit_db_0001_fFreeDBNoSelectQuery
End Sub
Function db_0001_fSqlOneFieldQuery(strSQL As String)
  Dim rst As Recordset
  Set rst = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
  rst.MoveFirst
  If rst.EOF Then
     db_0001_fSqlOneFieldQuery = CTE_db_no_record
  Else
     db_0001_fSqlOneFieldQuery = rst.Fields(0).Value
  End If
  rst.Close
End Function
Function db_0001_fDBRowExists(strSQL As String)
On Error GoTo Err_db_0001_fDBRowExists
  Dim rst As Recordset
  Set rst = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
  rst.MoveFirst
  rst.Close
Exit_db_0001_fDBRowExists_1:
  db_0001_fDBRowExists = 1
  Exit Function
Exit_db_0001_fDBRowExists_0:
  db_0001_fDBRowExists = 0
  Exit Function
Err_db_0001_fDBRowExists:
  rst.Close
  Resume Exit_db_0001_fDBRowExists_0
End Function
Function db_0001_fPrepareTxtFieldToBeInserted(ByVal linea As String)
  'Dos opciones: o usamos variable local ret, o usamos ByVal porque por defecto es ByRef
  db_0001_fPrepareTxtFieldToBeInserted = Replace(linea, "'", "''")
End Function
Function db_0001_fBuildSqlForCreateBigTableMemo(table_name As String, num_fields As Integer)
  Dim i As Integer
  Dim sql As String
  sql = ""
  sql = sql & "CREATE TABLE " & table_name & " ("
  For i = 1 To num_fields
    If i = 1 Then
      sql = sql & "field_" & i & " Memo"
    Else
      sql = sql & ", field_" & i & " Memo"
    End If
  Next i
  sql = sql & ")"
  db_0001_fBuildSqlForCreateBigTableMemo = sql
End Function
Function db_0001_fAllDataOfTwoDbTablesAreEqual(table_name1 As String, table_name2 As String) As Boolean
  Dim rst1 As Recordset
  Dim rst2 As Recordset
  Dim cont_f As Integer
  Dim cont_c As Integer
  Dim iguales As Boolean
  Set rst1 = CurrentDb.OpenRecordset("SELECT * FROM " & table_name1, dbOpenDynaset)
  Set rst2 = CurrentDb.OpenRecordset("SELECT * FROM " & table_name2, dbOpenDynaset)
  rst1.MoveFirst
  rst2.MoveFirst
  iguales = True
  For cont_c = 1 To rst1.Fields.Count
    If rst1.Fields(cont_c - 1).Name = rst2.Fields(cont_c - 1).Name Then
    Else
      Debug.Print "(" & cont_f & "," & cont_c & ") Los nombres de las columnas no coinciden: " & rst1.Fields(cont_c - 1).Name & " " & rst2.Fields(cont_c - 1).Name
      iguales = False
    End If
  Next cont_c
  cont_f = 0
  Do While (Not rst1.EOF And Not rst2.EOF And iguales = True)
    cont_f = cont_f + 1
    For cont_c = 1 To rst1.Fields.Count
      If rst1.Fields(cont_c - 1).Value = rst2.Fields(cont_c - 1).Value Then
      Else
        Debug.Print "(" & cont_f & "," & cont_c & ") Los valores de las columnas no coinciden: " & rst1.Fields(cont_c - 1).Value & " " & rst2.Fields(cont_c - 1).Value
        iguales = False
      End If
      DoEvents
    Next cont_c
    rst1.MoveNext
    rst2.MoveNext
    DoEvents
  Loop
  If iguales = True Then
    Debug.Print "Son iguales"
  Else
    Debug.Print "No son iguales"
  End If
  db_0001_fAllDataOfTwoDbTablesAreEqual = iguales
End Function
Function db_0001_fFreeDBNRowQuery(strSQL As String, ARR_rows As Variant)
  Dim ret As Boolean
  Dim i As Integer
  ret = True
  Dim rst As Recordset
  Set rst = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
  rst.MoveLast
  rst.MoveFirst
  ReDim ARR_rows(1 To rst.RecordCount) As Variant
  i = 0
  Do While Not rst.EOF
    i = i + 1
    ARR_rows(i) = rst.Fields(0).Value
    rst.MoveNext
    DoEvents
  Loop
  db_0001_fFreeDBNRowQuery = ret
End Function
Function db_0001_fFreeDBNRowNotArrayQuery(strSQL As String, outputFormat As Integer, return_if_empty As Variant) As Variant
  Dim ret As String
  Dim i As Integer
  ret = ""
  Dim rst As Recordset
  Set rst = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
  If rst.EOF Then
    ret = return_if_empty
  Else
    rst.MoveLast
    rst.MoveFirst
    i = 0
    Do While Not rst.EOF
      i = i + 1
      Select Case outputFormat
      Case CTE_db_output_format_4
        If i = 1 Then
          ret = ret & rst.Fields(0).Value
        Else
          ret = ret & ", " & rst.Fields(0).Value
        End If
      Case CTE_db_output_format_14
        If i = 1 Then
          ret = ret & rst.Fields(0).Value
        Else
          ret = ret & "," & rst.Fields(0).Value
        End If
      Case CTE_db_output_format_15
        If i = 1 Then
          ret = ret & "'" & rst.Fields(0).Value & "'"
        Else
          ret = ret & ", '" & rst.Fields(0).Value & "'"
        End If
      Case Else
        If i = 1 Then
          ret = ret & "'" & rst.Fields(0).Value & "'"
        Else
          ret = ret & ", '" & rst.Fields(0).Value & "'"
        End If
      End Select
      rst.MoveNext
      DoEvents
    Loop
  End If
  db_0001_fFreeDBNRowNotArrayQuery = ret
End Function
Function db_0001_fSaveLog(msg As String)
  Dim sql As String
  sql = "INSERT INTO t_log (txt) VALUES ('" & db_0001_fPrepareTxtFieldToBeInserted(msg) & "')"
  db_0001_fFreeDBNoSelectQuery (sql)
End Function

