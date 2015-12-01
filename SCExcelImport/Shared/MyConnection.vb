
Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports System.Text
Imports MySql.Data.MySqlClient
Imports System.Configuration
Imports System.Web.Configuration


Public Class MyConnection

    'Remarks added. will check from myconnection project

#Region "Initalise"



    Private Const COMMAND_TIMEOUT As Int32 = 100

    Private _DbType As DbConstant.DbType = DbConstant.DbType.ODBC
    Private _connString As String = ""
    Private _trans As IDbTransaction
    Private _isolationLevel As IsolationLevel
    Private _conn As IDbConnection
    Private _cmdTimeout As Int32
    Private _commandBehavior As CommandBehavior
    Private _transCount As Integer

    Private LogFileName As String = "QryLog.txt"
    Private FileToWrite As System.IO.StreamWriter

    Public Property CreateLog As Boolean

#End Region


    Public Property DatabaseType() As DbConstant.DbType
        Get
            Return _DbType
        End Get
        Set(ByVal value As DbConstant.DbType)
            _DbType = value
        End Set
    End Property

    Public Property ConnectionString() As String
        Get
            Return _connString
        End Get
        Set(ByVal value As String)
            _connString = value
        End Set
    End Property


    Public Property TransIsolationLevel() As IsolationLevel
        Get
            Return _isolationLevel
        End Get
        Set(ByVal Value As IsolationLevel)
            _isolationLevel = Value
        End Set
    End Property

    Public Property CmdTimeout() As Int32
        Get
            If _cmdTimeout = 0 Then
                Return COMMAND_TIMEOUT
            End If
            Return _cmdTimeout
        End Get
        Set(ByVal Value As Int32)
            _cmdTimeout = Value
        End Set
    End Property



    Public Sub New(ByVal mdbFile As String, Optional ByVal password As String = "")
        DatabaseType = DbConstant.DbType.OLEDB
        ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbFile & ";Jet OLEDB:Database Password= " & password & " ;"
    End Sub


    Public Sub New(ByVal DbType As String, _
                   ByVal ServerName As String, _
                   ByVal DbName As String, _
                   ByVal UserName As String, _
                   ByVal Password As String, _
                   ByVal Trust As String)


        Select Case Trim(UCase(DbType))
            Case "MYSQL"
                If ConnectionString = "" Then
                    ConnectionString = "Server=" & ServerName & ";Database=" & DbName & ";Uid=" & UserName & ";Pwd=" & Password & ";Convert Zero Datetime=True"
                End If
                DatabaseType = DbConstant.DbType.MYSQL
            Case "SQL"
                If ConnectionString = "" Then
                    If Trust = "Y" Then
                        ConnectionString = "Data Source=" & ServerName & ";Initial Catalog=" & DbName & ";Integrated Security=SSPI;"
                    Else
                        ConnectionString = "Data Source=" & ServerName & ";Initial Catalog=" & DbName & ";Connection Timeout=1000;User Id=" & UserName & ";Password=" & Password & ";"
                    End If

                End If
                DatabaseType = DbConstant.DbType.SQL
                'Data Source=myServerAddress;Initial Catalog=myDataBase;User Id=myUsername;Password=myPassword;
            Case Else
        End Select

    End Sub

    Public Sub New(ByVal ConnectString As String, ByVal DbType As DbConstant.DbType)
        ConnectionString = ConnectString
        DatabaseType = DbType
    End Sub


#Region "Connection"

    Public Function GetConnection() As IDbConnection
        Dim Connection As IDbConnection = Nothing

        Try
            Select Case DatabaseType
                Case DbConstant.DbType.ODBC
                    Connection = New OdbcConnection(ConnectionString)
                Case DbConstant.DbType.MYSQL
                    Connection = New MySqlConnection(ConnectionString)
                Case DbConstant.DbType.OLEDB
                    Connection = New OleDbConnection(ConnectionString)
                Case DbConstant.DbType.SQL
                    Connection = New SqlConnection(ConnectionString)
                Case Else
                    Connection = New OdbcConnection(ConnectionString)
            End Select
        Catch ex As Exception
            Throw
        End Try

        Return Connection
    End Function

    Public Function VerifyConnection() As Boolean
        Dim Connection As IDbConnection = Nothing

        Try
            Connection = GetConnection()
            Connection.Open()
            If Connection Is Nothing Then
                Return False
            Else
                Connection.Close()
                Return True
            End If
        Catch ex As Exception
            Throw
        Finally
            If Not Connection Is Nothing Then
                Connection.Dispose()
            End If
        End Try
        Return False
    End Function

#End Region


#Region "Commnad"

    Public Function GetCommand() As IDbCommand
        Dim Command As IDbCommand = Nothing

        Try
            Select Case DatabaseType
                Case DbConstant.DbType.MYSQL
                    Command = New MySqlCommand
                Case DbConstant.DbType.ODBC
                    Command = New OdbcCommand()
                Case DbConstant.DbType.OLEDB
                    Command = New OleDbCommand()
                Case DbConstant.DbType.SQL
                    Command = New SqlCommand()
                Case Else
                    Command = New OdbcCommand()
            End Select
        Catch ex As Exception
            Throw
        End Try
        Return Command

    End Function

    Public Function GetCommand(ByVal cmdText As String, _
                               ByVal Connection As IDbConnection) As IDbCommand
        Dim Command As IDbCommand = Nothing

        Try
            Select Case DatabaseType
                Case DbConstant.DbType.MYSQL
                    Command = New MySqlCommand(cmdText, Connection)
                Case DbConstant.DbType.ODBC
                    Command = New OdbcCommand(cmdText, Connection)
                Case DbConstant.DbType.OLEDB
                    Command = New OleDbCommand(cmdText, Connection)
                Case DbConstant.DbType.SQL
                    Command = New SqlCommand(cmdText, Connection)
                Case Else
                    Command = New OdbcCommand(cmdText, Connection)
            End Select
        Catch ex As Exception
            Throw
        End Try

        Return Command
    End Function


    Public Function GetCommand(ByVal strCmdText As String, _
                               ByVal cmdType As CommandType, _
                               ByVal cmdTimeout As Integer, _
                               ByVal ParameterArray As DbConstant.ParamStruct()) As IDbCommand

        Dim cmd As IDbCommand = GetCommand()
        Dim i As Int16

        Try
            If Not ParameterArray Is Nothing Then
                For i = 0 To ParameterArray.Length - 1
                    Dim ps As DbConstant.ParamStruct = ParameterArray(i)
                    Dim pm As IDbDataParameter = GetParameter(ps.ParamName, ps.direction, ps.value, ps.DataType, ps.sourceColumn, ps.size, ps.RowVersion)
                    cmd.Parameters.Add(pm)
                Next i
            End If
            cmd.CommandType = cmdType
            cmd.CommandText = strCmdText
        Catch ex As Exception
            Throw
        End Try
        Return cmd

    End Function
#End Region


#Region "Parameters"

    Public Function GetParameter() As IDbDataParameter

        Select Case DatabaseType
            Case DbConstant.DbType.MYSQL
                Return New MySqlParameter
            Case DbConstant.DbType.ODBC
                Return New OdbcParameter
            Case DbConstant.DbType.SQL
                Return New SqlParameter
            Case DbConstant.DbType.OLEDB
                Return New OleDbParameter
        End Select

        Return New OdbcParameter
    End Function

    Public Function GetParameter(ByVal paramName As String, _
                                        ByVal paramDirection As ParameterDirection, _
                                        ByVal paramValue As Object, _
                                        ByVal paramtype As DbType, _
                                        ByVal sourceColumn As String, _
                                        ByVal size As Int16, _
                                        ByVal RowVersion As System.Data.DataRowVersion) As IDbDataParameter
        Dim param As IDbDataParameter = GetParameter()
        param.ParameterName = paramName
        param.DbType = paramtype
        If size > 0 Then
            'param.Size = size
        End If
        If Not paramValue Is Nothing Then
            param.Value = paramValue
        End If
        param.Direction = paramDirection
        If Not sourceColumn = "" Then
            param.SourceColumn = sourceColumn
        End If
        param.SourceVersion = RowVersion

        Return (param)
    End Function

#End Region


#Region "Adapter"


    Public Function GetAdapter(ByVal command As IDbCommand) As IDataAdapter
        Dim Adapter As IDataAdapter = Nothing

        Select Case DatabaseType
            Case DbConstant.DbType.MYSQL
                Adapter = New MySqlDataAdapter(command)
            Case DbConstant.DbType.ODBC
                Adapter = New OdbcDataAdapter(command)
            Case DbConstant.DbType.OLEDB
                Adapter = New OleDbDataAdapter(command)
            Case DbConstant.DbType.SQL
                Adapter = New SqlDataAdapter(command)
            Case Else
                Adapter = New OdbcDataAdapter(command)
        End Select

        Return Adapter
    End Function


    Public Function GetAdapter() As IDataAdapter
        Dim Adapter As IDataAdapter = Nothing

        Select Case DatabaseType
            Case DbConstant.DbType.MYSQL
                Adapter = New MySqlDataAdapter()
            Case DbConstant.DbType.ODBC
                Adapter = New OdbcDataAdapter()
            Case DbConstant.DbType.OLEDB
                Adapter = New OleDbDataAdapter()
            Case DbConstant.DbType.SQL
                Adapter = New SqlDataAdapter()
            Case Else
                Adapter = New OdbcDataAdapter()
        End Select

        Return Adapter
    End Function

#End Region


#Region "CommandBuilder"
    Public Function GetCommandBuilder() As Object
        Select Case DatabaseType
            Case DbConstant.DbType.MYSQL
                Return New MySqlCommandBuilder
            Case DbConstant.DbType.ODBC
                Return New OdbcCommandBuilder
            Case DbConstant.DbType.SQL
                Return New SqlCommandBuilder
            Case DbConstant.DbType.OLEDB
                Return New OleDbCommandBuilder
            Case Else
                Return New OdbcCommandBuilder
        End Select
        Return Nothing
    End Function

#End Region


#Region "Transactions"

    Private Function GetTransaction(ByVal conn As IDbConnection, ByVal transisolationLevel As IsolationLevel) As IDbTransaction
        Return conn.BeginTransaction(transisolationLevel)
    End Function

    Public Sub BeginTrans(ByVal transisolationLevel As IsolationLevel)
        If _transCount = 0 Then
            _conn = GetConnection()
            _conn.Open()
            _trans = GetTransaction(_conn, transisolationLevel)
        End If
        _transCount = _transCount + 1
    End Sub
    Public Sub BeginTrans()
        If _transCount = 0 Then
            _conn = GetConnection()
            _conn.Open()
            _trans = GetTransaction(_conn, IsolationLevel.ReadCommitted)
        End If
        _transCount = _transCount + 1
    End Sub

    Public Sub CommitTrans(Optional ByVal CloseConnection As Boolean = True)
        _transCount = _transCount - 1
        If _transCount = 0 Then
            _CommitTrans(CloseConnection)
        End If
    End Sub

    Private Sub _CommitTrans(ByVal CloseConnection As Boolean)
        _trans.Commit()
        DisposeTrans(CloseConnection)
    End Sub

    Public Sub AbortTrans()
        If IsInTransaction() Then
            'If _transCount <> 0 Then
            _trans.Rollback()
            _transCount = 0
            DisposeTrans(True)
            'End If
        End If
    End Sub
    Public Sub RollBackTrans()
        AbortTrans()
    End Sub

    Public Sub CloseConnection()
        If Not _conn Is Nothing Then
            _conn.Close()
            _conn.Dispose()
        End If
    End Sub

    Private Sub DisposeTrans(ByVal CloseConnection As Boolean)
        If CloseConnection Then
            If Not _conn Is Nothing Then
                _conn.Close()
                _conn.Dispose()
            End If
        End If
        _trans.Dispose()
    End Sub

    Public Function IsInTransaction() As Boolean

        If _transCount > 0 Then
            Return True
        End If
        Return False
        'The following code dose not work.
        'The _trans object dose not become nothing
        'Return (Not _trans Is Nothing)
    End Function

#End Region


#Region "Prepare All"

    ' This method is used by ExecDataSet, ExecScalar, ExecReader and ExecNonQuery. This is a common piece of 
    ' code called in these methods
    Private Sub PrepareAll(ByRef cmd As IDbCommand, ByRef conn As IDbConnection, ByVal strSQL As String, ByVal cmdType As CommandType, ByVal parameterArray As DbConstant.ParamStruct())
        ' If transaction has already been started
        If Not IsInTransaction() Then
            Try
                conn = GetConnection()
                cmd = GetCommand(strSQL, cmdType, CmdTimeout, parameterArray)
                cmd.Connection = conn
                conn.Open()
            Catch ex As Exception
                Throw
            End Try
        Else
            cmd = GetCommand(strSQL, cmdType, CmdTimeout, parameterArray)
            cmd.Connection = _conn
            cmd.Transaction = _trans
        End If
    End Sub

#End Region

    'To return a DataSet after running a SQL Statement
#Region "ExecDataSet"



    Public Function ExecDataSet(ByVal ds As DataSet, _
                                ByVal strSQL As String, _
                                ByVal cmdtype As CommandType) As DataSet
        ExecDataSet(ds, strSQL, cmdtype, Nothing)
        Return ds
    End Function

    Public Function ExecDataSet(ByVal strSQL As String, _
                                ByVal cmdtype As CommandType) As DataSet
        Return ExecDataSet(strSQL, cmdtype, Nothing)
    End Function

    Public Function ExecDataSet(ByVal strSQL As String, _
                                ByVal cmdtype As CommandType, _
                                ByVal parameterArray As DbConstant.ParamStruct()) As DataSet

        Dim ds As New DataSet("DataSet")
        ExecDataSet(ds, strSQL, cmdtype, parameterArray)
        Return ds

    End Function

    Public Function ExecDataSet(ByVal ds As DataSet, _
                           ByVal strSQL As String, _
                           ByVal cmdtype As CommandType, _
                           ByVal parameterArray As DbConstant.ParamStruct()) As DataSet
        If CreateLog Then
            System.IO.File.AppendAllText(LogFileName, strSQL)
        End If

        Dim da As IDbDataAdapter = Nothing
        Dim cmd As IDbCommand = Nothing
        Dim conn As IDbConnection = Nothing
        Try
            da = GetAdapter()
            PrepareAll(cmd, conn, strSQL, cmdtype, parameterArray)
            da.SelectCommand = cmd
            da.FillSchema(ds, SchemaType.Source)
            da.Fill(ds)
            ds.EnforceConstraints = False
            Return ds
        Catch ex As Exception
            Throw
        Finally
            If Not IsInTransaction() Then
                If Not IsNothing(conn) Then
                    conn.Close()
                    conn.Dispose()
                End If
            End If
            If Not IsNothing(cmd) Then
                cmd.Dispose()
            End If
            If Not IsNothing(da) Then
                CType(da, IDisposable).Dispose()
            End If
        End Try
    End Function


#End Region

#Region "SaveDataSet"

    ' This method saves data in a dataset with a single table and mandates the table name to be "Table".
    ' Operations on a single table are batched.
    Public Function SaveDataSet(ByVal ds As DataSet, _
                           ByVal insertSQL As String, _
                           ByVal deleteSQL As String, _
                           ByVal UpdateSQL As String, _
                           ByVal SelectSql As String, _
                           ByVal InsertparameterArray As DbConstant.ParamStruct(), _
                           ByVal DeleteparameterArray As DbConstant.ParamStruct(), _
                           ByVal UpdateparameterArray As DbConstant.ParamStruct(), _
                           ByVal TableName As String) As DataSet

        If CreateLog Then
            System.IO.File.AppendAllText(LogFileName, insertSQL & vbCrLf & deleteSQL & vbCrLf & UpdateSQL & vbCrLf & SelectSql & vbCrLf)
        End If

        Dim cn As IDbConnection = Nothing
        Dim da As IDbDataAdapter = Nothing

        Dim TbleIndex As Integer

        Try
            da = GetAdapter()
            If Not IsInTransaction() Then
                cn = GetConnection()
                If insertSQL <> "" Then
                    da.InsertCommand = GetCommand(insertSQL, CommandType.Text, CmdTimeout, InsertparameterArray)
                    da.InsertCommand.Connection = cn
                End If
                If UpdateSQL <> "" Then
                    da.UpdateCommand = GetCommand(UpdateSQL, CommandType.Text, CmdTimeout, UpdateparameterArray)
                    da.UpdateCommand.Connection = cn
                End If
                If deleteSQL <> "" Then
                    da.DeleteCommand = GetCommand(deleteSQL, CommandType.Text, CmdTimeout, DeleteparameterArray)
                    da.DeleteCommand.Connection = cn
                End If
                'If SelectSql <> "" Then
                '    da.SelectCommand = GetCommand(SelectSql, CommandType.Text, CmdTimeout, DeleteparameterArray)
                '    da.SelectCommand.Connection = cn
                'End If
                cn.Open()
            Else
                If insertSQL <> "" Then
                    da.InsertCommand = GetCommand(insertSQL, CommandType.Text, CmdTimeout, InsertparameterArray)
                    da.InsertCommand.Connection = _conn
                    da.InsertCommand.Transaction = _trans
                End If
                If UpdateSQL <> "" Then
                    da.UpdateCommand = GetCommand(UpdateSQL, CommandType.Text, CmdTimeout, UpdateparameterArray)
                    da.UpdateCommand.Connection = _conn
                    da.UpdateCommand.Transaction = _trans
                End If
                If deleteSQL <> "" Then
                    da.DeleteCommand = GetCommand(deleteSQL, CommandType.Text, CmdTimeout, DeleteparameterArray)
                    da.DeleteCommand.Connection = _conn
                    da.DeleteCommand.Transaction = _trans
                End If
                'If SelectSql <> "" Then
                '    da.SelectCommand = GetCommand(SelectSql, CommandType.Text, CmdTimeout, DeleteparameterArray)
                '    da.SelectCommand.Connection = _conn
                '    da.SelectCommand.Transaction = _trans
                'End If
            End If
            If TableName = "" Then
                TableName = ds.Tables(0).TableName
            End If
            TbleIndex = ds.Tables.IndexOf(TableName)
            ds.Tables(TbleIndex).TableName = "Table"
            da.Update(ds)
            ds.Tables(TbleIndex).TableName = TableName
        Catch ex As Exception
            ds.Tables(TbleIndex).TableName = TableName
            'GenericExceptionHandler(ex)
            Throw

        Finally
            If Not IsInTransaction() Then
                cn.Close()
                cn.Dispose()
            End If
            If insertSQL <> "" Then
                da.InsertCommand.Parameters.Clear()
                da.InsertCommand.Dispose()
            End If
            If UpdateSQL <> "" Then
                da.UpdateCommand.Parameters.Clear()
                da.UpdateCommand.Dispose()
            End If
            If deleteSQL <> "" Then
                da.DeleteCommand.Parameters.Clear()
                da.DeleteCommand.Dispose()
            End If
            CType(da, IDisposable).Dispose()
        End Try
    End Function

#End Region


#Region "Data Access Methods"



    Public Function ExecuteNonQuery(ByVal SqlText As String) As Integer

        If CreateLog Then
            System.IO.File.AppendAllText(LogFileName, SqlText & vbCrLf)
        End If

        Dim Command As IDbCommand = Nothing
        Dim Connection As IDbConnection = Nothing

        Try

            If IsInTransaction() Then
                Connection = _conn
                Command = GetCommand(SqlText, Connection)
                Command.Transaction = _trans
            Else
                Connection = GetConnection()
                Command = GetCommand(SqlText, Connection)
                Connection.Open()
            End If
            Return Command.ExecuteNonQuery()
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
            'GenericExceptionHandler(ex)
            Throw
            Return 0
        Finally
            If Not IsInTransaction() Then
                If Not Command Is Nothing Then
                    Command.Dispose()
                End If

                If Not Connection Is Nothing Then
                    Connection.Dispose()
                End If
            End If
        End Try

    End Function

    Public Function ExecuteReader(ByVal SqlText As String) As IDataReader

        If CreateLog Then
            System.IO.File.AppendAllText(LogFileName, SqlText & vbCrLf)
        End If

        Dim Command As IDbCommand
        Dim Connection As IDbConnection

        Try
            Connection = GetConnection()
            Command = GetCommand(SqlText, Connection)
            Connection.Open()
            Return Command.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
            'Use the command behavior to automatically close
            'the connection when the reader is closed. We need
            'to leave the connection open until the data is
            'retrieved from the data source.
        Catch ex As Exception
            'System.Diagnostics.Debug.WriteLine(ex.Message)
            Throw
            Return Nothing
        Finally
        End Try
    End Function


    Public Function ExecuteScalar(ByVal sqlText As String) As Object

        If CreateLog Then
            System.IO.File.AppendAllText(LogFileName, sqlText & vbCrLf)

        End If

        Dim Command As IDbCommand = Nothing
        Dim Connection As IDbConnection = Nothing

        Try
            Connection = GetConnection()
            Command = GetCommand(sqlText, Connection)
            Connection.Open()
            Return Command.ExecuteScalar
        Catch ex As Exception
            'System.Diagnostics.Debug.WriteLine(ex.Message)
            Throw
        Finally
            If Not Command Is Nothing Then
                Command.Dispose()
            End If
            If Not Connection Is Nothing Then
                Connection.Dispose()
            End If
        End Try

        Return Nothing

    End Function

    Public Function FillDataSet(ByVal sqlText As String) As DataSet
        Dim DataSet As DataSet = Nothing
        Dim Adapter As IDataAdapter = Nothing
        Dim Command As IDbCommand = Nothing
        Dim Connection As IDbConnection = Nothing

        Try
            Connection = GetConnection()
            Command = GetCommand(sqlText, Connection)
            Adapter = GetAdapter(Command)
            DataSet = New DataSet
            Adapter.Fill(DataSet)
            Return DataSet
        Catch ex As Exception
            'System.Diagnostics.Debug.WriteLine(ex.Message)
            'GenericExceptionHandler(ex)
            Throw
            Return Nothing
        Finally
            If Not Command Is Nothing Then
                Command.Dispose()
            End If
            If Not Connection Is Nothing Then
                Connection.Dispose()
            End If
        End Try

        Return Nothing

    End Function

    Public Function FillDataTable(ByVal sqlText As String) As DataTable
        Dim DataSet As DataSet = Nothing
        Dim Adapter As IDataAdapter = Nothing
        Dim Command As IDbCommand = Nothing
        Dim Connection As IDbConnection = Nothing

        Try
            Connection = GetConnection()
            Command = GetCommand(sqlText, Connection)
            Adapter = GetAdapter(Command)
            DataSet = New DataSet
            Adapter.Fill(DataSet)
            Return DataSet.Tables(0)
        Catch ex As Exception
            'System.Diagnostics.Debug.WriteLine(ex.Message)
            'GenericExceptionHandler(ex)
            Throw
            Return Nothing
        Finally
            If Not Command Is Nothing Then
                Command.Dispose()
            End If
            If Not Connection Is Nothing Then
                Connection.Dispose()
            End If
        End Try

        Return Nothing

    End Function
    Public Function GetRecordSet(ByVal SqlText As String, ByVal TableName As String) As DataTable
        Dim DataSet As DataSet = Nothing
        Dim Adapter As IDataAdapter = Nothing
        Dim Command As IDbCommand = Nothing
        Dim Connection As IDbConnection = Nothing
        Dim dtTable As DataTable = Nothing

        Try
            Connection = GetConnection()
            Command = GetCommand(SqlText, Connection)
            Adapter = GetAdapter(Command)
            DataSet = New DataSet
            Adapter.Fill(DataSet)
            DataSet.Tables(0).TableName = TableName
            dtTable = DataSet.Tables(0)
            DataSet.Tables.Remove(dtTable)
            Return dtTable
        Catch ex As Exception
            'System.Diagnostics.Debug.WriteLine(ex.Message)
            'GenericExceptionHandler(ex)
            Throw
            Return Nothing
        Finally
            If Not Command Is Nothing Then
                Command.Dispose()
            End If
            If Not Connection Is Nothing Then
                Connection.Dispose()
            End If
        End Try

        Return Nothing

    End Function

#End Region

#Region "Exception handlers"

    'Private Sub GenericExceptionHandler(ByVal ex As Exception)

    '    If TypeOf ex Is SqlException Then
    '        SQLExceptionHandler(ex)
    '    ElseIf TypeOf ex Is OleDbException Then
    '        OLEDBExceptionHandler(ex)
    '    ElseIf TypeOf ex Is OdbcException Then
    '        ODBCExceptionHandler(ex)
    '    ElseIf TypeOf ex Is MySqlException Then
    '        MySQLExceptionHandler(ex)
    '    Else
    '        Throw
    '    End If

    'End Sub

    'Private Sub MySQLExceptionHandler(ByVal ex As MySqlException)


    '    Dim sb As New StringBuilder
    '    'For Each sqlerr In ex.Errors
    '    '    sb.AppendFormat("Error: {0}{1}", sqlerr.Message, Environment.NewLine)
    '    '    sb.AppendFormat("Server: {0}{1}", sqlerr.Code, Environment.NewLine)
    '    '    sb.AppendFormat("Source: {0}{1}", sqlerr.Level, Environment.NewLine)
    '    '    sb.Append("-----------------------------------------------")
    '    'Next
    '    'TODO For each custom sql server error have an entry
    '    sb.AppendFormat("Error: {0}{1}", ex.Message, Environment.NewLine)
    '    sb.AppendFormat("Server: {0}{1}", ex.Number, Environment.NewLine)
    '    sb.AppendFormat("Source: {0}{1}", ex.ErrorCode, Environment.NewLine)
    '    sb.Append("-----------------------------------------------")
    '    Throw New Exception(sb.ToString, ex)

    'End Sub

    'Private Sub SQLExceptionHandler(ByVal ex As SqlException)
    '    Dim sqlerr As SqlError
    '    Dim sb As New StringBuilder
    '    For Each sqlerr In ex.Errors
    '        sb.AppendFormat("Error: {0}{1}", sqlerr.Message, Environment.NewLine)
    '        sb.AppendFormat("Server: {0}{1}", sqlerr.Server, Environment.NewLine)
    '        sb.AppendFormat("Source: {0}{1}", sqlerr.Source, Environment.NewLine)
    '        sb.Append("-----------------------------------------------")
    '    Next
    '    'TODO For each custom sql server error have an entry
    '    Throw New Exception(sb.ToString, ex)
    'End Sub

    'Private Sub OLEDBExceptionHandler(ByVal ex As OleDbException)
    '    Dim oledberr As OleDbError
    '    Dim sb As New StringBuilder
    '    For Each oledberr In ex.Errors
    '        sb.AppendFormat("Error: {0}{1}", oledberr.Message, Environment.NewLine)
    '        sb.AppendFormat("Source: {0}{1}", oledberr.Source, Environment.NewLine)
    '        sb.Append("-----------------------------------------------")
    '    Next
    '    'TODO For each custom sql server error have an entry
    '    Throw New Exception(sb.ToString, ex)
    'End Sub

    'Private Sub ODBCExceptionHandler(ByVal ex As OdbcException)
    '    Dim odbcerr As OdbcError
    '    Dim sb As New StringBuilder
    '    For Each odbcerr In ex.Errors
    '        sb.AppendFormat("Error: {0}{1}", odbcerr.Message, Environment.NewLine)
    '        sb.AppendFormat("Source: {0}{1}", odbcerr.Source, Environment.NewLine)
    '        sb.Append("-----------------------------------------------")
    '    Next
    '    'TODO For each custom sql server error have an entry
    '    Throw New Exception(sb.ToString, ex)
    'End Sub

#End Region


#Region "Misc"

    Public Function CheckTableExists(ByVal TableName As String) As Boolean

        Dim sql As String
        Dim ds As DataSet

        sql = ""
        sql = sql & "  Select * From Information_Schema.Tables"
        sql = sql & "  Where       Table_Catalog='" & getDbName() & "'"
        sql = sql & "          And Table_Type = 'BASE TABLE'"
        sql = sql & "          And Table_Name='" & TableName & "'"
        ds = ExecDataSet(sql, CommandType.Text)
        If ds.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function getDbName() As String
        Dim sql As String
        Dim ds As DataSet

        sql = "SELECT DB_NAME() AS DataBaseName"
        ds = ExecDataSet(sql, CommandType.Text)
        If ds.Tables(0).Rows.Count > 0 Then
            Return ds.Tables(0).Rows(0).Item(0)
        Else
            Return ""
        End If
    End Function

#End Region

End Class




