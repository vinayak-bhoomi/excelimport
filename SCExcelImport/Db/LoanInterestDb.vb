Public Class LoanInterestDb
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
            sql = sql & " DELETE FROM " & tbl.bnkLoaninterest & " "
            sql = sql & " WHERE GLCode IN ('" & glcodes & "')"
            objCon.ExecuteNonQuery(sql)

        End If

        sql = ""
        sql = sql & " SELECT * "
        sql = sql & " FROM " & tbl.bnkLoaninterest & " "
        sql = sql & " WHERE GLCode IN ('" & glcodes & "')"
        ds = objCon.ExecDataSet(sql, CommandType.Text)

        ds.Tables(0).TableName = tbl.bnkLoaninterest
        ds.Tables(0).PrimaryKey = New DataColumn() {ds.Tables(0).Columns("GLCODE"), ds.Tables(0).Columns("AccountNo"),
                                                    ds.Tables(0).Columns("LoanSerialNo"), ds.Tables(0).Columns("ApplicableDate")}
        dt = ds.Tables(0)

    End Sub

    Public Function GetUpdatedDs() As DataSet
        Return ds
    End Function

    Public Sub Add(ByVal oLimit As LoanInterest)

        dr = dt.Select("GLCode='" & oLimit.Glcode & "' " & _
                        " AND AccountNo=" & oLimit.AccountNo & "" &
                        " AND LoanSerialNo = " & oLimit.LoanSerialNo & "").FirstOrDefault

        If IsNothing(dr) Then
            dr = dt.Rows.Add
        End If
        dr.Item("GlCode") = oLimit.Glcode
        dr.Item("AccountNo") = oLimit.AccountNo
        dr.Item("ApplicableDate") = oLimit.ApplicableDate
        dr.Item("Rate") = oLimit.Rate
        dr.Item("LoanSerialNo") = oLimit.LoanSerialNo
        'dr.Item("isLog") = "Y"
        dr.Item("UptoAmount") = 0
        dr.Item("Rate2") = 0
        dr.Item("UptoAmount2") = 0
        dr.Item("Rate3") = 0
        dr.Item("UptoAmount3") = 0

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
        pk.Add("LoanSerialNo")
        pk.Add("ApplicableDate")
        BuildCommand(tbl.bnkLoaninterest, ds.Tables(0), pk)

        For Each dr In ds.Tables(0).Rows
            If dr.RowState = DataRowState.Unchanged Then
            Else
                RemoveNull(dr)
            End If
        Next

        objCon.BeginTrans(IsolationLevel.ReadCommitted)

        Try
            objCon.SaveDataSet(ds, InsertCommand, DeleteCommand, UpdateCommand, SelectCommand, pmInsert, pmDelete, pmUpdate, tbl.bnkLoaninterest)
        Catch ex As Exception
            objCon.AbortTrans()
            Throw ex
        End Try

        objCon.CommitTrans(False)

    End Sub

End Class

Public Class LoanInterest

    Public Property Glcode As String
    Public Property InstallmentType As String
    Public Property AccountNo As Long
    Public Property LoanSerialNo As Integer
    Public Property ApplicableDate As Date
    Public Property Rate As Double

End Class

