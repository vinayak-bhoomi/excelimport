Public Class DepositReceiptDb
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
            sql = sql & " DELETE FROM " & tbl.bnkDepositReceipt & " "
            objCon.ExecuteNonQuery(sql)

        End If

        sql = ""
        sql = sql & " SELECT * "
        sql = sql & " FROM " & tbl.bnkDepositReceipt & " "
        sql = sql & " WHERE 1=0 "
        ds = objCon.ExecDataSet(sql, CommandType.Text)
        ds.Tables(0).TableName = tbl.bnkDepositReceipt

        dt = ds.Tables(0)
    End Sub

    Public Function GetUpdatedDs() As DataSet
        Return ds
    End Function

    Public Sub Add(ByVal oFD As DepositReceipt)

        dr = dt.NewRow
        dr.Item("DepositReceiptID") = ""
        dr.Item("GlCode") = oFD.GLCode
        dr.Item("AccountNo") = oFD.AccountNo
        dr.Item("ReceiptNo") = oFD.ReceiptNo
        dr.Item("Period") = oFD.Period
        dr.Item("PeriodCategory") = oFD.PeriodCategory

        dr.Item("DepositDate") = DBNull.Value
        If IsValidDate(oFD.DepositDate) Then
            dr.Item("DepositDate") = oFD.DepositDate
        End If

        dr.Item("RenewalDate") = DBNull.Value
        If IsValidDate(oFD.RenewalDate) Then
            dr.Item("RenewalDate") = oFD.RenewalDate
        End If
        dr.Item("DepositAmount") = oFD.DepositAmount

        dr.Item("MaturityDate") = DBNull.Value
        If IsValidDate(oFD.MaturityDate) Then
            dr.Item("MaturityDate") = oFD.MaturityDate
        End If
        dr.Item("RateOfInterest") = oFD.RateOfInterest
        dr.Item("MaturityAmount") = oFD.MaturityAmount
        dr.Item("MBNPDate") = DBNull.Value

        dr.Item("PaidDate") = DBNull.Value
        dr.Item("ReceiptStatus") = "A"
        If IsValidDate(oFD.PaidDate) Then
            dr.Item("PaidDate") = oFD.PaidDate
            dr.Item("ReceiptStatus") = "P"
        End If
        dr.Item("InterestAccrued") = 0
        dr.Item("InterestPaid") = 0
        dr.Item("ReceiptMarked") = ""
        dr.Item("PledgedGlCode") = ""
        dr.Item("PledgedAccountNo") = 0
        dr.Item("InterestGlCode") = ""
        dr.Item("InterestAccountNo") = 0
        dr.Item("PrintedNoOfCopies") = 0
        dr.Item("UserID") = ""
        dr.Item("AuthenticateUserID") = ""
        dr.Item("DateTimeCreation") = DateTime.Now
        dr.Item("ReserveFundID") = ""
        'dr.Item("isLog") = "Y"
        dr.Item("EPMode") = ""
        dr.Item("EPBankCode") = ""
        dr.Item("EPAccountType") = ""
        dr.Item("EPAccountNo") = 0
        dr.Item("BranchCode") = oFD.BranchCode
        dr.Item("IntCategory") = ""
        dr.Item("ReceiptMode") = ""
        dr.Item("RISRate") = 0
        dr.Item("BookCode") = ""
        dr.Item("VoucherNo") = 0
        dr.Item("VoucherDate") = DBNull.Value
        dr.Item("IssuedDate") = DBNull.Value
        dr.Item("PrevReceiptNo") = 0
        dr.Item("RBIBankSrNo") = 0
        dt.Rows.Add(dr)

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
        pk.Add("ReceiptNo")
        BuildCommand(tbl.bnkDepositReceipt, ds.Tables(0), pk)

        For Each dr In ds.Tables(0).Rows
            If dr.RowState = DataRowState.Unchanged Then
            Else
                RemoveNull(dr)
            End If
        Next

        objCon.BeginTrans(IsolationLevel.ReadCommitted)

        Try
            objCon.SaveDataSet(ds, InsertCommand, DeleteCommand, UpdateCommand, SelectCommand, pmInsert, pmDelete, pmUpdate, tbl.bnkDepositReceipt)
        Catch ex As Exception
            objCon.AbortTrans()
            Throw ex
        End Try

        objCon.CommitTrans(False)

    End Sub


End Class

Public Class DepositReceipt

    Public Property GLCode As String
    Public Property AccountNo As Long
    Public Property ReceiptNo As String
    Public Property Period As String
    Public Property PeriodCategory As String
    Public Property DepositDate As Date
    Public Property RenewalDate As Date
    Public Property DepositAmount As Double
    Public Property MaturityDate As Date
    Public Property RateOfInterest As Double
    Public Property MaturityAmount As Double
    Public Property MBNPDate As Date
    Public Property PaidDate As Date
    Public Property InterestAccrued As Double
    Public Property InterestPaid As Double
    Public Property ReceiptStatus As String
    Public Property ReceiptMarked As String
    Public Property PledgedGlCode As String
    Public Property PledgedAccountNo As Long
    Public Property InterestGlCode As String
    Public Property InterestAccountNo As Long
    Public Property PrintedNoOfCopies As Long
    Public Property UserID As String
    Public Property AuthenticateUserID As String
    Public Property DateTimeCreation As DateTime
    Public Property ReserveFundID As String
    Public Property isLog As String
    Public Property EPMode As String
    Public Property EPBankCode As String
    Public Property EPAccountType As String
    Public Property EPAccountNo As Long
    Public Property BranchCode As String
    Public Property IntCategory As String
    Public Property ReceiptMode As String
    Public Property RISRate As Long
    Public Property BookCode As String
    Public Property VoucherNo As Long
    Public Property VoucherDate As Date
    Public Property IssuedDate As Date
    Public Property PrevReceiptNo As Long
    Public Property RBIBankSrNo As Long

End Class

