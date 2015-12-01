Public Class AddressDb

    Inherits SaveDB

    Private ds As DataSet
    Private dt As DataTable
    Private dr As DataRow

    Private tbl As New Tables

    Public Sub New(ByVal _conn As MyConnection)
        _objCon = _conn
    End Sub

    Public Sub Init(ByVal emptyDb As Boolean)

        Dim sql As String

        If emptyDb Then
            sql = ""
            sql = sql & " DELETE FROM " & tbl.mstAddress & " "
            _objCon.ExecuteNonQuery(sql)
        End If

        sql = ""
        sql = sql & " SELECT * "
        sql = sql & " FROM " & tbl.mstAddress & " "

        sql = sql & " WHERE 1=0 "
        ds = _objCon.ExecDataSet(sql, CommandType.Text)
        ds.Tables(0).TableName = tbl.mstAddress

        dt = ds.Tables(0)
    End Sub

    Public Sub BeginTran()

        Dim sql As String

        sql = ""
        sql = sql & " SELECT * "
        sql = sql & " FROM " & tbl.mstAddress & " "
        ds = _objCon.ExecDataSet(sql, CommandType.Text)
        ds.Tables(0).TableName = tbl.mstAddress
        ds.Tables(0).PrimaryKey = New DataColumn() {ds.Tables(0).Columns("AddressId")}

        dt = ds.Tables(0)
    End Sub

    Public Function GetUpdatedDs() As DataSet
        Return ds
    End Function

    Public Sub Add(ByVal oAdd As Address)

        dr = dt.Select("AddressId='" & oAdd.AddressId & "'").FirstOrDefault
        If IsNothing(dr) Then
            dr = dt.Rows.Add
        End If
        dr.Item("AddressId") = oAdd.AddressId
        dr.Item("SrNo") = oAdd.SrNo
        dr.Item("TypeOfAddress") = oAdd.TypeOfAddress
        dr.Item("Address1") = Trim(oAdd.Address1)
        dr.Item("Address2") = Trim(oAdd.Address2)
        dr.Item("Address3") = Trim(oAdd.Address3)
        dr.Item("IsAddressForCommunication") = ""
        dr.Item("City") = Trim(oAdd.City)
        dr.Item("State") = Trim(oAdd.State)
        dr.Item("Pincode") = Trim(oAdd.Pincode)
        dr.Item("TelephoneNo") = Trim(oAdd.TelephoneNo)
        dr.Item("Address1_Language") = ""
        dr.Item("Address2_Language") = ""
        dr.Item("Address3_Language") = ""
        dr.Item("City_Language") = ""
        dr.Item("State_Language") = ""

    End Sub

    Public Sub Add(ByVal addId As Long, ByVal fullAdd As String)

        Dim oAdd As Address
        Dim lstAdd As List(Of String)

        lstAdd = refineStr(fullAdd)

        oAdd = New Address
        oAdd.AddressId = addId
        oAdd.SrNo = 1
        Select Case lstAdd.Count
            Case 1
                oAdd.Address1 = lstAdd(0)
            Case 2
                oAdd.Address1 = lstAdd(0)
                oAdd.Address2 = lstAdd(1)
            Case 3
                oAdd.Address1 = lstAdd(0)
                oAdd.Address2 = lstAdd(1)
                oAdd.Address3 = lstAdd(2)
            Case 4
                oAdd.Address1 = lstAdd(0)
                oAdd.Address2 = lstAdd(1)
                oAdd.Address3 = lstAdd(2)
                oAdd.City = lstAdd(3)
            Case 5
                oAdd.Address1 = JoinStr(",", lstAdd(0), lstAdd(1))
                oAdd.Address2 = lstAdd(2)
                oAdd.Address3 = lstAdd(3)
                oAdd.City = lstAdd(4)
            Case 6
                oAdd.Address1 = JoinStr(",", lstAdd(0), lstAdd(1))
                oAdd.Address2 = JoinStr(",", lstAdd(2), lstAdd(3))
                oAdd.Address3 = lstAdd(4)
                oAdd.City = lstAdd(5)
        End Select

        Add(oAdd)

    End Sub

    Private Function refineStr(ByVal strValue As String) As List(Of String)

        Dim str As String
        Dim arryStr As List(Of String)
        Dim reslt As New List(Of String)

        arryStr = strValue.Split(",").ToList

        For i As Integer = 0 To arryStr.Count - 1
            str = Trim(arryStr(i))
            If str <> "" Then
                reslt.Add(str)
            End If
        Next
        Return reslt
    End Function

    Public Sub Save()

        Dim pk As ArrayList
        Dim dr As DataRow

        If ds.Tables.Count <= 0 Then
            Return
        End If

        pk = New ArrayList()
        pk.Add("AddressId")
        BuildCommand(tbl.mstAddress, ds.Tables(0), pk)

        For Each dr In ds.Tables(0).Rows
            If dr.RowState = DataRowState.Unchanged Then
            Else
                RemoveNull(dr)
            End If
        Next

        _objCon.BeginTrans(IsolationLevel.ReadCommitted)

        Try
            _objCon.SaveDataSet(ds, InsertCommand, DeleteCommand, UpdateCommand, SelectCommand, pmInsert, pmDelete, pmUpdate, tbl.mstAddress)
        Catch ex As Exception
            _objCon.AbortTrans()
            Throw ex
        End Try

        _objCon.CommitTrans(False)

    End Sub

    Public Function IsValidDate(ByVal dtObj As Object)
        If IsNothing(dtObj) Then
            Return False
        End If

        If Not IsDate(dtObj) Then
            Return False
        End If

        If dtObj = Date.MinValue Then
            Return False
        End If

        If dtObj <= CDate("1899-12-30") Then
            Return False
        End If

        Return True
    End Function

End Class

Public Class Address

    Public Property AddressId As String
    Public Property SrNo As Integer
    Public Property TypeOfAddress As String
    Public Property Address1 As String
    Public Property Address2 As String
    Public Property Address3 As String
    Public Property IsAddressForCommunication As String
    Public Property City As String
    Public Property State As String
    Public Property Pincode As String
    Public Property TelephoneNo As String
    Public Property Address1_Language As String
    Public Property Address2_Language As String
    Public Property Address3_Language As String
    Public Property City_Language As String
    Public Property State_Language As String

    Public Sub New()
        AddressId = ""
        SrNo = 0
        TypeOfAddress = ""
        Address1 = ""
        Address2 = ""
        Address3 = ""
        IsAddressForCommunication = ""
        City = ""
        State = ""
        Pincode = ""
        TelephoneNo = ""
        Address1_Language = ""
        Address2_Language = ""
        Address3_Language = ""
        City_Language = ""
        State_Language = ""
    End Sub
End Class