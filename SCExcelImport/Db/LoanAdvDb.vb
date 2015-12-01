Public Class LoanAdvDb
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
            sql = sql & " DELETE FROM " & tbl.bnkLoanAdvances & " "
            sql = sql & " WHERE GLCode IN ('" & glcodes & "')"
            objCon.ExecuteNonQuery(sql)

        End If

        sql = ""
        sql = sql & " SELECT * "
        sql = sql & " FROM " & tbl.bnkLoanAdvances & " "
        sql = sql & " WHERE GLCode IN ('" & glcodes & "')"
        ds = objCon.ExecDataSet(sql, CommandType.Text)

        ds.Tables(0).TableName = tbl.bnkLoanAdvances
        ds.Tables(0).PrimaryKey = New DataColumn() {ds.Tables(0).Columns("GLCODE"), ds.Tables(0).Columns("AccountNo"),
                                                    ds.Tables(0).Columns("LoanSerialNo"), ds.Tables(0).Columns("ApplicableDate")}
        dt = ds.Tables(0)

    End Sub

    Public Function GetUpdatedDs() As DataSet
        Return ds
    End Function

    Public Sub Add(ByVal oLimit As LoanAdv)

        dr = dt.Select("GLCode='" & oLimit.Glcode & "' " & _
                        " AND AccountNo=" & oLimit.AccountNo & "").FirstOrDefault

        If IsNothing(dr) Then
            dr = dt.Rows.Add
        End If
        dr.Item("LoanAdvanceID") = ""
        dr.Item("GLCode") = oLimit.Glcode
        dr.Item("AccountNo") = oLimit.AccountNo
        dr.Item("PurposeMasterID") = ""
        dr.Item("SecterMasterID") = ""
        dr.Item("NPAStatus") = 1
        dr.Item("PrioritySecterMasterID") = ""
        dr.Item("SecurityMasterID") = ""
        dr.Item("SecurityAmount") = 0
        dr.Item("SanctionedAmount") = 0
        dr.Item("SanctionedDate") = DBNull.Value
        dr.Item("SecurityRemarks") = ""
        dr.Item("DirectorsRelative") = 0
        dr.Item("DirectorName") = ""
        dr.Item("DirectorCode") = 0
        dr.Item("Remarks") = ""
        dr.Item("CompanyName") = ""
        dr.Item("CompanyAddress1") = ""
        dr.Item("CompanyAddress2") = ""
        dr.Item("CompanyAddress3") = ""
        dr.Item("EmployeeNo") = ""
        dr.Item("Department") = ""
        dr.Item("Paysheet") = ""
        dr.Item("UserID") = ""
        dr.Item("DateTimeCreation") = DateTime.Now
        dr.Item("AuthenticateUserID") = ""
        dr.Item("IsSecured") = "N"
        dr.Item("RelationShip") = ""
        dr.Item("SuitSectionCode") = ""
        dr.Item("isDisableInterestPosting") = "N"
        dr.Item("isSuitFiled") = "N"
        'dr.Item("isLog") = "Y"
        dr.Item("isDisableInterestCalculation") = "N"
        dr.Item("SuitSectionCode2") = ""
        dr.Item("isDisableLoanNotice") = "N"
        dr.Item("isDisablePenalPosting") = "N"
        dr.Item("isDisableChargesPosting") = "N"
        dr.Item("StopNoticeDate") = DBNull.Value
        dr.Item("StopLoanIntDate") = DBNull.Value
        dr.Item("CaseFileDate") = DBNull.Value

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
        BuildCommand(tbl.bnkLoanAdvances, ds.Tables(0), pk)

        For Each dr In ds.Tables(0).Rows
            If dr.RowState = DataRowState.Unchanged Then
            Else
                RemoveNull(dr)
            End If
        Next

        objCon.BeginTrans(IsolationLevel.ReadCommitted)

        Try
            objCon.SaveDataSet(ds, InsertCommand, DeleteCommand, UpdateCommand, SelectCommand, pmInsert, pmDelete, pmUpdate, tbl.bnkLoanAdvances)
        Catch ex As Exception
            objCon.AbortTrans()
            Throw ex
        End Try

        objCon.CommitTrans(False)

    End Sub

End Class

Public Class LoanAdv

    Public Property LoanAdvanceID As String
    Public Property Glcode As String
    Public Property AccountNo As Long

End Class

