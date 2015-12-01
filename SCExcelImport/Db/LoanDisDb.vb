Public Class LoanDisDb
    Inherits Bhoomi.Migration.DB.SaveDB

    Private ds As DataSet
    Private dt As DataTable
    Private dr As DataRow

    Private tbl As New Bhoomi.Migration.Shared.Tables

    Public Sub New(ByVal _conn As Bhoomi.Migration.DB.MyConnection)
        objCon = _conn
    End Sub

    Public Sub BeginTran(ByVal emptyDb As Boolean, ByVal glcodes As String)

        Dim sql As String

        glcodes = glcodes.Replace(",", "','")
        If emptyDb Then

            sql = ""
            sql = sql & " DELETE FROM " & tbl.bnkDisbursement & " "
            sql = sql & " WHERE GLCode IN ('" & glcodes & "')"
            objCon.ExecuteNonQuery(sql)

        End If

        sql = ""
        sql = sql & " SELECT * "
        sql = sql & " FROM " & tbl.bnkDisbursement & " "
        sql = sql & " WHERE GLCode IN ('" & glcodes & "')"
        ds = objCon.ExecDataSet(sql, CommandType.Text)

        ds.Tables(0).TableName = tbl.bnkDisbursement
        ds.Tables(0).PrimaryKey = New DataColumn() {ds.Tables(0).Columns("GLCODE"), ds.Tables(0).Columns("AccountNumber"),
                                                    ds.Tables(0).Columns("LoanSerialNo")}
        dt = ds.Tables(0)

    End Sub

    Public Function GetUpdatedDs() As DataSet
        Return ds
    End Function

    Public Sub Add(ByVal oLimit As LoanDisbursment)

        dr = dt.Select("GLCode='" & oLimit.Glcode & "' " & _
                        " AND AccountNumber=" & oLimit.AccountNumber & " " &
                        " AND LoanSerialNo=" & oLimit.LoanSerialNo & "").FirstOrDefault

        If IsNothing(dr) Then
            dr = dt.Rows.Add
        End If

        dr.Item("DisbursementID") = ""
        dr.Item("DisbursementDate") = oLimit.DisbursementDate
        dr.Item("Amount") = oLimit.Amount
        dr.Item("TransactionNo") = ""
        dr.Item("UserID") = ""
        dr.Item("DateTimeCreation") = DateTime.Now
        dr.Item("LoanLimitId") = ""
        dr.Item("GLcode") = oLimit.Glcode
        dr.Item("AccountNumber") = oLimit.AccountNumber
        dr.Item("LoanSerialNo") = 1
        'dr.Item("isLog") = "Y"
        dr.Item("DeductShares") = 0
        dr.Item("DeductLoans") = 0
        dr.Item("DeductOthers") = 0

    End Sub

    Public Sub Save()

        Dim pk As ArrayList
        Dim dr As DataRow

        If ds.Tables.Count <= 0 Then
            Return
        End If

        pk = New ArrayList()
        pk.Add("GLCode")
        pk.Add("AccountNumber")
        BuildCommand(tbl.bnkDisbursement, ds.Tables(0), pk)

        For Each dr In ds.Tables(0).Rows
            If dr.RowState = DataRowState.Unchanged Then
            Else
                RemoveNull(dr)
            End If
        Next

        objCon.BeginTrans(IsolationLevel.ReadCommitted)

        Try
            objCon.SaveDataSet(ds, InsertCommand, DeleteCommand, UpdateCommand, SelectCommand, pmInsert, pmDelete, pmUpdate, tbl.bnkDisbursement)
        Catch ex As Exception
            objCon.AbortTrans()
            Throw ex
        End Try

        objCon.CommitTrans(False)

    End Sub

End Class

Public Class LoanDisbursment

    Public Property TransactionNo As String
    Public Property Glcode As String
    Public Property AccountNumber As Long
    Public Property DisbursementDate As Date
    Public Property Amount As Double
    Public Property LoanSerialNo As Integer

End Class
