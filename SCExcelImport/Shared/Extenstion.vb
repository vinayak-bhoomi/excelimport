Imports System.Runtime.CompilerServices

Public Module Extenstions

    <Extension()>
    Public Function HasColumn(dr As IDataReader, columnName As String) As Boolean
        For Each row As DataRow In dr.GetSchemaTable().Rows
            If row("ColumnName").ToString() = columnName Then
                Return True
            End If
        Next
        Return False
    End Function

    <Extension()>
    Public Function GetSafeVal(ByVal dr As DataRow, ByVal field As String) As Object

        If Not IsDBNull(dr.Item(field)) Then
            Return dr.Item(field)
        End If

        'return empty value
        Select Case dr.Table.Columns(field).DataType
            Case Type.GetType("System.Double"), Type.GetType("System.Int16"), Type.GetType("System.Int32"), Type.GetType("System.Int64")
                Return 0
            Case Type.GetType("system.string")
                Return ""
            Case Else
                Return ""
        End Select

    End Function

    <Extension()>
    Public Function GetSafeDbl(ByVal dr As DataRow, ByVal field As String) As Object

        If Not IsDBNull(dr.Item(field)) Then
            Return Convert.ToDouble(Val(dr.Item(field)))
        End If

        Return 0

        'return empty value
        'Select Case dr.Table.Columns(field).DataType
        '    Case Type.GetType("System.Double"), Type.GetType("System.Int16"), Type.GetType("System.Int32"), Type.GetType("System.Int64")
        '        Return 0
        '    Case Type.GetType("system.string")
        '        Return ""
        '    Case Else
        '        Return ""
        'End Select

    End Function

    <Extension()>
    Public Function GetSafeVal(ByVal dre As IDataRecord, ByVal field As String)
        If Not IsDBNull(dre.Item(field)) Then
            Return dre.Item(field)
        End If
        Return dre.Item(field)
    End Function

    <Extension()>
    Public Function IsLastRow(ByVal dr As DataRow) As Boolean

        Dim currentIndex As Integer
        Dim recordCount As Integer

        currentIndex = dr.Table.Rows.IndexOf(dr) + 1
        recordCount = dr.Table.Rows.Count()

        If currentIndex = recordCount Then
            Return True
        End If
        Return False

    End Function

    <Extension()>
    Public Function IsFirstRow(ByVal dr As DataRow) As Boolean

        Dim currentIndex As Integer

        currentIndex = dr.Table.Rows.IndexOf(dr) + 1
        If currentIndex = 1 Then
            Return True
        End If

        Return False

    End Function

    <Extension()>
    Public Function Index(ByVal dr As DataRow) As Integer

        Dim _currentIndex As Integer

        _currentIndex = dr.Table.Rows.IndexOf(dr) + 1
        Return _currentIndex

    End Function

    <Extension()>
    Public Function IsNextRowChange(ByVal dr As DataRow, ByVal field As String) As Boolean

        If dr.IsLastRow Then
            Return True
        End If

        Dim currentIndex As Integer

        currentIndex = dr.Table.Rows.IndexOf(dr)

        If dr.Item(field) <> dr.Table.Rows(currentIndex + 1).Item(field) Then
            Return True
        End If

        Return False

    End Function

    <Extension()>
    Public Function IsBrek(ByVal dr As DataRow, ByVal interval As Integer) As Boolean

        Dim currentIndex As Integer
        Dim recordCount As Integer

        currentIndex = dr.Table.Rows.IndexOf(dr) + 1
        recordCount = dr.Table.Rows.Count()

        'Check is last row
        If currentIndex = recordCount Then
            Return True
        End If

        Dim result As Decimal
        result = currentIndex / interval

        If result.ToString.IndexOf(".") = -1 Then
            Return True
        End If

        Return False
    End Function

End Module
