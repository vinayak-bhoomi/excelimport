Imports System.Data

Public Class SaveDB


    Public SelectCommand As String = ""
    Public InsertCommand As String = ""
    Public UpdateCommand As String = ""
    Public DeleteCommand As String = ""

    Public _objCon As MyConnection

    Public pmInsert() As DbConstant.ParamStruct
    Public pmUpdate() As DbConstant.ParamStruct
    Public pmDelete() As DbConstant.ParamStruct

    Public Overridable Function insertPara(ByRef Para() As DbConstant.ParamStruct, _
                                           ByRef Index As Integer, _
                                           ByRef Direction As ParameterDirection, _
                                           ByRef SourceColumn As String, _
                                           ByRef ParaType As DbType, _
                                           ByRef ParaSize As Integer, _
                                           ByRef RowVersion As DataRowVersion) As DbConstant.ParamStruct()

        Para(Index).direction = Direction
        If RowVersion = DataRowVersion.Original Then
            Para(Index).ParamName = "@Original_" & SourceColumn
        Else
            Para(Index).ParamName = "@" & SourceColumn
        End If

        Para(Index).DataType = ParaType
        Para(Index).sourceColumn = SourceColumn
        Para(Index).size = 50
        Para(Index).RowVersion = RowVersion

        Return Para

    End Function

    Public Overridable Function BuildCommand(ByVal TableName As String, _
                                 ByVal dt As DataTable, _
                                 ByVal Pk As ArrayList) As Boolean

        GenerateCommand(TableName, dt, Pk)
        BuildParameter(TableName, dt, Pk)

        Return True
    End Function


    Public Overridable Function BuildParameter(ByVal TableName As String, _
                                   ByVal dt As DataTable, _
                                   ByVal Pk As ArrayList) As Boolean


        Dim dc As DataColumn
        Dim i As Integer = 0
        Dim r As Integer

        Dim InsertIndex As Integer = dt.Columns.Count - 1
        Dim UpdateIndex As Integer = dt.Columns.Count + Pk.Count - 1
        Dim DeleteIndex As Integer = Pk.Count - 1

        For Each dc In dt.Columns
            If dc.AutoIncrement Then
                InsertIndex = InsertIndex - 1
                UpdateIndex = UpdateIndex - 1
            ElseIf dc.DataType.ToString.ToUpper = "SYSTEM.BYTE[]" Then
                InsertIndex = InsertIndex - 1
                UpdateIndex = UpdateIndex - 1
            End If
        Next

        pmInsert = New DbConstant.ParamStruct(InsertIndex) {}
        pmUpdate = New DbConstant.ParamStruct(UpdateIndex) {}
        pmDelete = New DbConstant.ParamStruct(DeleteIndex) {}

        InsertIndex = 0
        UpdateIndex = 0
        DeleteIndex = 0
        For Each dc In dt.Columns
            If Not dc.AutoIncrement Then
                Select Case dc.DataType.ToString.ToUpper
                    Case "SYSTEM.STRING"
                        insertPara(pmInsert, InsertIndex, ParameterDirection.Input, dc.ColumnName, DbType.String, dc.MaxLength, DataRowVersion.Current)
                        insertPara(pmUpdate, UpdateIndex, ParameterDirection.Input, dc.ColumnName, DbType.String, dc.MaxLength, DataRowVersion.Current)
                        InsertIndex = InsertIndex + 1
                        UpdateIndex = UpdateIndex + 1
                    Case "SYSTEM.DATETIME"
                        insertPara(pmInsert, InsertIndex, ParameterDirection.Input, dc.ColumnName, DbType.DateTime, dc.MaxLength, DataRowVersion.Current)
                        insertPara(pmUpdate, UpdateIndex, ParameterDirection.Input, dc.ColumnName, DbType.DateTime, dc.MaxLength, DataRowVersion.Current)
                        InsertIndex = InsertIndex + 1
                        UpdateIndex = UpdateIndex + 1
                    Case "SYSTEM.GUID"
                        insertPara(pmInsert, InsertIndex, ParameterDirection.Input, dc.ColumnName, DbType.Guid, dc.MaxLength, DataRowVersion.Current)
                        insertPara(pmUpdate, UpdateIndex, ParameterDirection.Input, dc.ColumnName, DbType.Guid, dc.MaxLength, DataRowVersion.Current)
                        InsertIndex = InsertIndex + 1
                        UpdateIndex = UpdateIndex + 1
                    Case "SYSTEM.BYTE[]"
                    Case Else
                        insertPara(pmInsert, InsertIndex, ParameterDirection.Input, dc.ColumnName, DbType.String, dc.MaxLength, DataRowVersion.Current)
                        insertPara(pmUpdate, UpdateIndex, ParameterDirection.Input, dc.ColumnName, DbType.String, dc.MaxLength, DataRowVersion.Current)
                        InsertIndex = InsertIndex + 1
                        UpdateIndex = UpdateIndex + 1
                End Select
            End If
        Next

        i = 0
        For r = 0 To Pk.Count - 1
            For Each dc In dt.Columns
                If dc.ColumnName.ToUpper = Pk.Item(r).ToString.ToUpper Then
                    Select Case dc.DataType.ToString.ToUpper
                        Case "SYSTEM.STRING"
                            insertPara(pmDelete, DeleteIndex, ParameterDirection.Input, dc.ColumnName, DbType.String, dc.MaxLength, DataRowVersion.Current)
                            DeleteIndex = DeleteIndex + 1

                            insertPara(pmUpdate, UpdateIndex, ParameterDirection.Input, dc.ColumnName, DbType.String, dc.MaxLength, DataRowVersion.Original)
                            UpdateIndex = UpdateIndex + 1
                        Case "SYSTEM.GUID"
                            insertPara(pmDelete, DeleteIndex, ParameterDirection.Input, dc.ColumnName, DbType.String, dc.MaxLength, DataRowVersion.Current)
                            DeleteIndex = DeleteIndex + 1

                            insertPara(pmUpdate, UpdateIndex, ParameterDirection.Input, dc.ColumnName, DbType.String, dc.MaxLength, DataRowVersion.Original)
                            UpdateIndex = UpdateIndex + 1

                        Case "SYSTEM.DATETIME"
                            insertPara(pmDelete, DeleteIndex, ParameterDirection.Input, dc.ColumnName, DbType.DateTime, dc.MaxLength, DataRowVersion.Current)
                            DeleteIndex = DeleteIndex + 1

                            insertPara(pmUpdate, UpdateIndex, ParameterDirection.Input, dc.ColumnName, DbType.DateTime, dc.MaxLength, DataRowVersion.Original)
                            UpdateIndex = UpdateIndex + 1

                        Case "SYSTEM.BYTE[]"
                        Case Else
                            insertPara(pmDelete, DeleteIndex, ParameterDirection.Input, dc.ColumnName, DbType.String, dc.MaxLength, DataRowVersion.Current)
                            DeleteIndex = DeleteIndex + 1

                            insertPara(pmUpdate, UpdateIndex, ParameterDirection.Input, dc.ColumnName, DbType.String, dc.MaxLength, DataRowVersion.Original)
                            UpdateIndex = UpdateIndex + 1

                    End Select
                End If
            Next
        Next
        If DeleteIndex <> Pk.Count Then
            DeleteCommand = ""
        End If

        Return True


    End Function


    Public Overridable Function GenerateCommand(ByVal TableName As String, _
                                    ByVal dt As DataTable, _
                                    ByVal Pk As ArrayList) As Boolean

        Dim Commnad(3) As String
        Dim dc As DataColumn
        Dim i As Integer = 0

        Dim ColumnCount As Integer = 0

        If Pk.Count = 0 Then
            InsertCommand = ""
            UpdateCommand = ""
            DeleteCommand = ""
            SelectCommand = ""
            Return True
        End If
        ColumnCount = dt.Columns.Count
        For Each dc In dt.Columns
            If dc.AutoIncrement Then
                ColumnCount = ColumnCount - 1
                Exit For
            End If
        Next

        InsertCommand = " Insert into " & TableName & " ("
        DeleteCommand = " Delete From " & TableName
        SelectCommand = " Select * From " & TableName
        UpdateCommand = " Update " & TableName & " Set "

        Dim insertValues As String = ""
        i = 0
        For Each dc In dt.Columns

            If dc.AutoIncrement Then
            ElseIf dc.DataType.ToString.ToUpper = "SYSTEM.BYTE[]" Then
                i = i + 1
            Else
                i = i + 1
                If ColumnCount = i Then
                    InsertCommand = InsertCommand & dc.ColumnName & " "
                    insertValues = insertValues & "@" & dc.ColumnName & " "
                    UpdateCommand = UpdateCommand & dc.ColumnName & "=@" & dc.ColumnName & " "
                Else
                    InsertCommand = InsertCommand & dc.ColumnName & ","
                    insertValues = insertValues & "@" & dc.ColumnName & ","
                    UpdateCommand = UpdateCommand & dc.ColumnName & "=@" & dc.ColumnName & ","
                End If
            End If
        Next

        InsertCommand = InsertCommand & ") Values (" & insertValues & ") "
        UpdateCommand = UpdateCommand & " Where "
        DeleteCommand = DeleteCommand & " Where "
        SelectCommand = SelectCommand & " Where "
        For i = 0 To Pk.Count - 1
            If i = Pk.Count - 1 Then
                UpdateCommand = UpdateCommand & Pk.Item(i).ToString & "=@Original_" & Pk.Item(i).ToString
                DeleteCommand = DeleteCommand & Pk.Item(i).ToString & "=@" & Pk.Item(i).ToString
            Else
                UpdateCommand = UpdateCommand & Pk.Item(i).ToString & "=@Original_" & Pk.Item(i).ToString & " And "
                DeleteCommand = DeleteCommand & Pk.Item(i).ToString & "=@" & Pk.Item(i).ToString & " And "
            End If
            If i = 0 Then
                SelectCommand = SelectCommand & Pk.Item(i).ToString & "=SCOPE_IDENTITY();" ' & Pk.Item(i).ToString
            End If
        Next

        Return True

    End Function

    Public Sub RemoveNull(ByVal dr As DataRow)

        For Each dc In dr.Table.Columns
            Select Case dc.datatype.ToString.ToUpper
                Case "SYSTEM.STRING"
                    If IsDBNull(dr.Item(dc)) Then
                        dr.Item(dc) = ""
                    End If
                Case "SYSTEM.DATETIME"
                Case "SYSTEM.DECIMAL", "SYSTEM.DOUBLE", "SYSTEM.INT16", "SYSTEM.INT32", "SYSTEM.INT64"
                    If IsDBNull(dr.Item(dc)) Then
                        dr.Item(dc) = 0
                    End If
                Case "SYSTEM.BYTE[]"
                Case "SYSTEM.GUID"
                    If IsDBNull(dr.Item(dc)) Then
                        dr.Item(dc) = Guid.Empty
                    Else
                    End If
                Case "SYSTEM.BOOLEAN"
                Case Else
                    MsgBox("REMOVE NULL ERROR")
            End Select
        Next

    End Sub



End Class
