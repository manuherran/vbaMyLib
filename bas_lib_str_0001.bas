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
' str_0001_fLPad
' str_0001_fStringMultilineToOneLineTrim
' str_0001_fRemoveAllQuotes
' str_0001_fRemoveAllSpacesAndTabs
' str_0001_leftSideIs
' str_0001_rightSideIs
' str_0001_leftSideOf
' str_0001_rightSideOf
' str_0001_firstCapital
' 
' -----------------------------------------------------------------------
Function str_0001_fLPad(txt As String, pad As String, max_len As Integer)
  Dim ret As String
  ret = txt
  While Len(ret) < max_len
    ret = pad & ret
  Wend
  str_0001_fLPad = ret
End Function
Function str_0001_fStringMultilineToOneLineTrim(txt As String)
  Dim ret As String
  ret = txt
  ret = Replace(ret, vbTab, " ")
  ret = Replace(ret, vbCrLf, " ")
  Do While (InStr(ret, "  "))
    ret = Replace(ret, "  ", " ")
    DoEvents
  Loop
  ret = Trim(ret)
  str_0001_fStringMultilineToOneLineTrim = ret
End Function
Function str_0001_fRemoveAllQuotes(txt As String)
  Dim ret As String
  ret = txt
  ret = Replace(ret, "'", "")
  ret = Replace(ret, """", "")
  str_0001_fRemoveAllQuotes = ret
End Function
Function str_0001_fRemoveAllSpacesAndTabs(txt As String)
  Dim ret As String
  ret = txt
  ret = Replace(ret, " ", "")
  ret = Replace(ret, vbTab, "")
  ret = Trim(ret)
  str_0001_fRemoveAllSpacesAndTabs = ret
End Function
Function str_0001_leftSideIs(stringText As String, leftSideText As String)
  If (Left(stringText, Len(leftSideText)) = leftSideText) Then
    str_0001_leftSideIs = True
  Else
    str_0001_leftSideIs = False
  End If
End Function
Function str_0001_rightSideIs(stringText As String, rightSideText As String)
  If (Right(stringText, Len(rightSideText)) = rightSideText) Then
    str_0001_rightSideIs = True
  Else
    str_0001_rightSideIs = False
  End If
End Function
Function str_0001_leftSideOf(stringText As String, textToSearch As String)
  Dim pos As Long
  pos = InStr(stringText, textToSearch)
  str_0001_leftSideOf = Left(stringText, pos - 1)
End Function
Function str_0001_rightSideOf(stringText As String, textToSearch As String)
  Dim pos As Long
  Dim strLength As Long
  pos = InStr(stringText, textToSearch)
  strLength = Len(stringText)
  If pos = 0 Or pos = strLength Then
    str_0001_rightSideOf = ""
  Else
    str_0001_rightSideOf = Right(stringText, strLength - pos - Len(textToSearch) + 1)
  End If
End Function
Function str_0001_firstCapital(stringText As String)
  str_0001_firstCapital = UCase(Left(stringText, 1)) & LCase(Right(stringText, Len(stringText) - 1))
End Function
