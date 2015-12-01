
Public Class LoanLimitDb
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
            sql = sql & " DELETE FROM " & tbl.bnkLoanLimit & " "
            sql = sql & " WHERE GLCode IN ('" & glcodes & "')"
            objCon.ExecuteNonQuery(sql)

        End If

        sql = ""
        sql = sql & " SELECT * "
        sql = sql & " FROM " & tbl.bnkLoanLimit & " "
        sql = sql & " WHERE GLCode IN ('" & glcodes & "')"
        ds = objCon.ExecDataSet(sql, CommandType.Text)

        ds.Tables(0).TableName = tbl.bnkLoanLimit
        ds.Tables(0).PrimaryKey = New DataColumn() {ds.Tables(0).Columns("GLCODE"), ds.Tables(0).Columns("AccountNo"), ds.Tables(0).Columns("LoanSerialNo")}
        dt = ds.Tables(0)

    End Sub

    Public Function GetUpdatedDs() As DataSet
        Return ds
    End Function

    Public Sub Add(ByVal oLimit As LoanLimit)

        dr = dt.Select("GLCode='" & oLimit.glcode & "' AND AccountNo=" & oLimit.accountno & "").FirstOrDefault
        If IsNothing(dr) Then
            dr = dt.Rows.Add
        End If
        dr.Item("DateTimeCreation") = DateTime.Now
        dr.Item("LimitSanctioned") = oLimit.LimitSanctioned
        dr.Item("InstallmentAmount") = oLimit.InstallmentAmount
        dr.Item("SanctionedDate") = DBNull.Value
        If IsValidDate(oLimit.SanctionedDate) Then
            dr.Item("SanctionedDate") = oLimit.SanctionedDate
        End If
        dr.Item("MaturityDate") = DBNull.Value
        If IsValidDate(oLimit.MaturityDate) Then
            dr.Item("MaturityDate") = oLimit.MaturityDate
        End If
        dr.Item("FirstInstallmentDate") = oLimit.FirstInstallmentDate
        dr.Item("NoOfInstallment") = oLimit.NoOfInstallment
        dr.Item("InstallmentFrequency") = "M"
        dr.Item("RateOfInterest") = oLimit.RateOfInterest
        dr.Item("MoratoriumPeriod") = 0
        dr.Item("Margin") = ""
        dr.Item("LimitCategory") = ""
        dr.Item("DisbursementID") = ""
        dr.Item("UserID") = ""
        dr.Item("LoanAdvanceId") = ""
        dr.Item("LimitStatus") = oLimit.LimitStatus
        dr.Item("ClosedOn") = DBNull.Value
        If IsValidDate(oLimit.ClosedOn) Then
            dr.Item("ClosedOn") = oLimit.ClosedOn
        End If
        dr.Item("glcode") = oLimit.glcode
        dr.Item("InstallmentType") = oLimit.InstallmentType
        dr.Item("accountno") = oLimit.accountno
        dr.Item("InsuranceType") = DBNull.Value
        If oLimit.InsuranceType <> "" Then
            dr.Item("InsuranceType") = oLimit.InsuranceType
        End If
        dr.Item("ClosedBy") = DBNull.Value
        If oLimit.ClosedBy <> "" Then
            dr.Item("ClosedBy") = oLimit.ClosedBy
        End If
        dr.Item("InterestAmount") = 0
        dr.Item("LoanSerialNo") = 1
        dr.Item("FormNo") = 0
        dr.Item("FormDate") = DBNull.Value
        dr.Item("BondNo") = 0
        dr.Item("BondDate") = DBNull.Value
        'dr.Item("isLog") = "Y"
        dr.Item("BranchCode") = oLimit.BranchCode
        dr.Item("PurposeCode") = oLimit.PurposeCode

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
        BuildCommand(tbl.bnkLoanLimit, ds.Tables(0), pk)

        For Each dr In ds.Tables(0).Rows
            If dr.RowState = DataRowState.Unchanged Then
            Else
                RemoveNull(dr)
            End If
        Next

        objCon.BeginTrans(IsolationLevel.ReadCommitted)

        Try
            objCon.SaveDataSet(ds, InsertCommand, DeleteCommand, UpdateCommand, SelectCommand, pmInsert, pmDelete, pmUpdate, tbl.bnkLoanLimit)
        Catch ex As Exception
            objCon.AbortTrans()
            Throw ex
        End Try

        objCon.CommitTrans(False)

    End Sub

End Class

Public Class LoanLimit

    Public Property DateTimeCreation As Date
    Public Property LimitSanctioned As Double
    Public Property InstallmentAmount As Double
    Public Property SanctionedDate As Date
    Public Property MaturityDate As Date
    Public Property FirstInstallmentDate As Date
    Public Property NoOfInstallment As Integer
    Public Property InstallmentFrequency As String
    Public Property RateOfInterest As Double
    Public Property MoratoriumPeriod As Integer
    Public Property Margin As String
    Public Property LimitCategory As String
    Public Property DisbursementID As String
    Public Property UserID As String
    Public Property LoanAdvanceId As String
    Public Property LimitStatus As String
    Public Property ClosedOn As Date
    Public Property glcode As String
    Public Property InstallmentType As String
    Public Property accountno As Long
    Public Property InsuranceType As String
    Public Property ClosedBy As String
    Public Property InterestAmount As Integer
    Public Property LoanSerialNo As Integer
    Public Property FormNo As Integer
    Public Property FormDate As Date
    Public Property BondNo As Integer
    Public Property BondDate As Date
    Public Property isLog As String
    Public Property BranchCode As String
    Public Property PurposeCode As String

End Class

