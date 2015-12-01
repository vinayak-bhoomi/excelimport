Public Class DbConstant

    <Serializable()> _
Public Structure ParamStruct
        Public ParamName As String
        Public DataType As DbType
        Public value As Object
        Public direction As ParameterDirection
        Public sourceColumn As String
        Public size As Int16
        Public RowVersion As System.Data.DataRowVersion
    End Structure

    Public Enum DbType
        SQL = 1
        MYSQL = 2
        OLEDB = 3
        ODBC = 4
    End Enum

End Class
