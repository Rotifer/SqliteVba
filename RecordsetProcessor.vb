Option Explicit
' Wrapper class for Recordset class
Private m_recset As ADODB.Recordset
Private m_columnNames() As String
Private m_fieldsIndexColumnNamesMap As Dictionary
Private m_rows As Variant
Public Function Setup(recset As ADODB.Recordset) As Boolean
    Set m_recset = recset
    SetFieldsIndexColumnNamesMap
    SetRows
    Setup = True
End Function
Private Function SetFieldsIndexColumnNamesMap() As Boolean
    Dim i As Integer
    Dim fieldsIndexColumnNamesMap As Dictionary
    Dim columnName As String
    Set fieldsIndexColumnNamesMap = New Dictionary
    For i = 0 To m_recset.Fields.count - 1
        columnName = m_recset.Fields(i).Name
        fieldsIndexColumnNamesMap.Add i, columnName
    Next i
    Set m_fieldsIndexColumnNamesMap = fieldsIndexColumnNamesMap
    SetFieldsIndexColumnNamesMap = True
End Function
Public Function SetRows() As Boolean
    Dim rows As Variant
    rows = m_recset.GetRows
    m_rows = rows
    SetRows = True
End Function
Public Function GetColumnNames() As Variant
    GetColumnNames = m_fieldsIndexColumnNamesMap.Items
End Function
Private Function GetIndexForColumnName(columnName As String) As Integer
    Dim key As Variant
    For Each key In m_fieldsIndexColumnNamesMap.Keys
        If m_fieldsIndexColumnNamesMap(key) = columnName Then
            GetIndexForColumnName = key
            Exit Function
        End If
    Next key
    GetIndexForColumnName = -1
End Function
Public Function GetValuesForColumn(columnName As String) As Variant()
    Dim columnIndex As Integer
    Dim i As Integer
    Dim valuesForColumn() As Variant
    If Not m_fieldsIndexColumnNamesMap.Exists(columnName) Then
        GetValuesForColumn = valuesForColumn
        Exit Function
    End If
    columnIndex = GetIndexForColumnName(columnName)
    For i = 0 To UBound(m_rows, 2)
        ReDim Preserve valuesForColumn(i)
        valuesForColumn(i) = m_rows(columnIndex, i)
    Next i
    GetValuesForColumn = valuesForColumn
End Function
