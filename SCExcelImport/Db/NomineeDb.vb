Public Class NomineeDb
    Inherits Bhoomi.Migration.DB.SaveDB

    Private ds As DataSet
    Private dt As DataTable
    Private dr As DataRow

    Private tbl As New Bhoomi.Migration.Shared.Tables

    Public Sub New(ByVal _conn As Bhoomi.Migration.DB.MyConnection)
        objCon = _conn
    End Sub

    Public Sub BeginTran(ByVal emptyDb As Boolean)

        Dim sql As String

        If emptyDb Then

            sql = ""
            sql = sql & " DELETE FROM " & tbl.mstNominee & " "
            objCon.ExecuteNonQuery(sql)

        End If

        sql = ""
        sql = sql & " SELECT * "
        sql = sql & " FROM " & tbl.mstNominee & " "
        ds = objCon.ExecDataSet(sql, CommandType.Text)

        ds.Tables(0).TableName = tbl.mstNominee
        ds.Tables(0).PrimaryKey = New DataColumn() {ds.Tables(0).Columns("GLCode"), ds.Tables(0).Columns("Accountno"), ds.Tables(0).Columns("SrNo")}
        dt = ds.Tables(0)

    End Sub

    Public Function GetUpdatedDs() As DataSet
        Return ds
    End Function

    Public Sub Add(ByVal oNom As Nominee)

        dr = dt.Select("GLCode='" & oNom.Glcode & "' AND AccountNo=" & oNom.Accountno & " AND SrNo=" & oNom.SrNo & "").FirstOrDefault

        If IsNothing(dr) Then
            dr = dt.Rows.Add
        End If

        dr.Item("Glcode") = oNom.Glcode
        dr.Item("Accountno") = oNom.Accountno
        dr.Item("ReceiptNo") = oNom.ReceiptNo
        dr.Item("SrNo") = 1
        dr.Item("Name") = oNom.Name
        dr.Item("NameinLanguage") = ""
        dr.Item("AddressId") = oNom.AddressId
        dr.Item("CustomerNo") = oNom.CustomerNo
        dr.Item("CustomerType") = ""
        dr.Item("Glcode1") = ""
        dr.Item("Accountno1") = 0
        dr.Item("Relation") = oNom.Relation
        dr.Item("NomineeRegNo") = ""
        dr.Item("RegDate") = DBNull.Value
        dr.Item("isActive") = "Y"
        dr.Item("Age") = oNom.Age
        dr.Item("IsMinor") = "N"
        dr.Item("Percentage") = "100"
        dr.Item("GaurdianName") = ""
        dr.Item("RelationShipWithGaurdian") = ""
        dr.Item("GaurdianAddressID") = ""
        dr.Item("Userid") = ""
        dr.Item("DateTimeCreation") = Now.Date
        dr.Item("MinorDOB") = DBNull.Value

    End Sub

    Public Sub Save()

        Dim pk As ArrayList
        Dim dr As DataRow

        If ds.Tables.Count <= 0 Then
            Return
        End If

        pk = New ArrayList()
        pk.Add("GLCode")
        pk.Add("AccountNo")
        pk.Add("SrNo")

        BuildCommand(tbl.mstNominee, ds.Tables(0), pk)

        For Each dr In ds.Tables(0).Rows
            If dr.RowState = DataRowState.Unchanged Then
            Else
                RemoveNull(dr)
            End If
        Next

        objCon.BeginTrans(IsolationLevel.ReadCommitted)

        Try
            objCon.SaveDataSet(ds, InsertCommand, DeleteCommand, UpdateCommand, SelectCommand, pmInsert, pmDelete, pmUpdate, tbl.mstNominee)
        Catch ex As Exception
            objCon.AbortTrans()
            Throw ex
        End Try

        objCon.CommitTrans(False)

    End Sub

End Class

Public Class Nominee

    Public Property Glcode As String
    Public Property Accountno As Long
    Public Property ReceiptNo As String
    Public Property SrNo As Integer
    Public Property Name As String
    Public Property NameinLanguage As String
    Public Property AddressId As String
    Public Property CustomerNo As Long
    Public Property CustomerType As String
    Public Property Glcode1 As String
    Public Property Relation As String
    Public Property NomineeRegNo As String
    Public Property RegDate As Date
    Public Property Age As Integer
    Public Property IsMinor As String

End Class
