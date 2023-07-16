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
' array_0001_fCreateAsoc1DArray
' array_0001_fAsoc1DArrayAddValue
' array_0001_fTestAsoc1DArray
' 
' 
' 
' 
' 
' 
' -----------------------------------------------------------------------
Function array_0001_fCreateAsoc1DArray(AR_data() As Variant, AR_keys() As Variant) As Collection
  Dim i As Long
  Dim cont As Long
  Dim myCollection As Collection
  Set myCollection = New Collection
  cont = 0
  For i = LBound(AR_data()) To UBound(AR_data())
    cont = cont + 1
    myCollection.Add AR_data(cont), AR_keys(cont)
  Next i
  Set array_0001_fCreateAsoc1DArray = myCollection
End Function
Sub array_0001_fAsoc1DArrayAddValue(p_data As Variant, p_keys As Variant, myCollection As Collection)
  myCollection.Add p_data, p_keys
End Sub
Sub array_0001_fTestAsoc1DArray()
  ReDim AR_data(1 To 2) As Variant
  ReDim AR_keys(1 To 2) As Variant
  Dim AAR_myAsocArray As Collection
  Set AAR_myAsocArray = New Collection
  AR_data(1) = "dato1"
  AR_data(2) = "dato2"
  AR_keys(1) = "key1"
  AR_keys(2) = "key2"
  Set AAR_myAsocArray = array_0001_fCreateAsoc1DArray(AR_data(), AR_keys())
  Debug.Print AAR_myAsocArray("key1")
  Debug.Print AAR_myAsocArray("key2")
  'Debug.Print AAR_myAsocArray("key3") ERROR
  ReDim Preserve AR_data(1 To 3) As Variant
  ReDim Preserve AR_keys(1 To 3) As Variant
  Debug.Print AAR_myAsocArray("key1")
  Debug.Print AAR_myAsocArray("key2")
  array_0001_fAsoc1DArrayAddValue "dato3", "key3", AAR_myAsocArray
  Debug.Print AAR_myAsocArray("key3")
End Sub






Function array_0001_fCreateAsoc2DArray() As Collection
  Dim i As Long
  Dim cont As Long
  Dim myCollection As Collection
  Set myCollection = New Collection
  cont = 0
  For i = LBound(AR_data()) To UBound(AR_data())
    cont = cont + 1
    myCollection.Add AR_data(cont), AR_keys(cont)
  Next i
  Set array_0001_fCreateAsoc2DArray = myCollection
End Function
Sub array_0001_fAsoc2DArrayAddValue(p_data As Variant, p_keys1 As Variant, p_keys2 As Variant, myCollection As Collection)
  
  'myCollection.Add p_data, p_keys1
  
  'Dim myCollection2 As Collection
  'Set myCollection2 = New Collection
  
  myCollection.Add AddCollectionItem(), p_keys1
  myCollection(p_keys1).Add AddCollectionItem(), p_keys2
  
  myCollection(p_keys1)(p_keys2).Add p_data
End Sub
Sub array_0001_fTestAsoc2DArray()
  ReDim AR_data(1 To 2) As Variant
  ReDim AR_keys1(1 To 2) As Variant
  ReDim AR_keys2(1 To 2) As Variant
  Dim AAR_myAsocArray As Collection
  Set AAR_myAsocArray = New Collection


  array_0001_fAsoc2DArrayAddValue "dato1A", "key1", "keyA", AAR_myAsocArray
  array_0001_fAsoc2DArrayAddValue "dato1B", "key1", "keyB", AAR_myAsocArray
  array_0001_fAsoc2DArrayAddValue "dato2A", "key2", "keyA", AAR_myAsocArray
  array_0001_fAsoc2DArrayAddValue "dato2B", "key2", "keyB", AAR_myAsocArray

  'Set AAR_myAsocArray = array_0001_fCreateAsoc2DArray(AR_data(), AR_keys1(), AR_keys2())
  Debug.Print AAR_myAsocArray("key1")("keyA")
  Debug.Print AAR_myAsocArray("key2")("keyB")
  Debug.Print AAR_myAsocArray("key3")("keyC"); Error
End Sub
Public Function AddCollectionItem() As Collection
  Set AddCollectionItem = New Collection
End Function
