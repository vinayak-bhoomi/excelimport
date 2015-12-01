
Public Class TransactionDb
    Inherits SaveDB

    Private ds As DataSet
    Private dt As DataTable
    Private dr As DataRow

    Private tbl As New Tables

    Public Sub New(ByVal _conn As MyConnection)
        _objCon = _conn
    End Sub

    Public Sub Init(Optional ByVal emptyDb As Boolean = False)

        Dim sql As String

        If emptyDb Then

            sql = ""
            sql = sql & " DELETE FROM " & tbl.bnk_MainTransaction & " "
            _objCon.ExecuteNonQuery(sql)

        End If

        sql = ""
        sql = sql & " SELECT * "
        sql = sql & " FROM " & tbl.bnk_MainTransaction & " "
        ds = _objCon.ExecDataSet(sql, CommandType.Text)

        ds.Tables(0).TableName = tbl.bnk_MainTransaction
        ds.Tables(0).PrimaryKey = New DataColumn() {ds.Tables(0).Columns("DailyTransactionId"), ds.Tables(0).Columns("Date")}
        dt = ds.Tables(0)
    End Sub

    Public Sub Add(ByVal oTrn As Transaction)

        dr = dt.Rows.Add
        dr("DailyTransactionId") = String.Format("{0:0000000}", CInt(oTrn.DailyTransactionId))
        dr("VoucherNo") = oTrn.VoucherNo
        dr("SerialNumber") = 0
        dr("glcode") = oTrn.GLCode
        dr("Date") = oTrn.TrDate
        dr("AccountNumber") = oTrn.AccountNumber
        dr("customerNo") = oTrn.CustomerNo
        dr("BookCode") = oTrn.BookCode
        dr("BookType") = "X"
        dr("BranchCode") = oTrn.BranchCode
        dr("PresentDate") = DBNull.Value
        dr("DueDate") = DBNull.Value
        dr("DbCr") = oTrn.DrCr
        dr("Narration") = oTrn.Narration
        dr("Amount") = oTrn.Amount
        dr("InstrumentAmount") = 0
        dr("NarrationinMarathi") = ""
        dr("InstrumentNumber") = "0"
        dr("InstrumentDate") = DBNull.Value 'Format(m_Date, "YYYY-MM-DD 00:00:00")
        dr("instrumenttype") = ""
        dr("SlipNo") = 0
        dr("TokenNumber") = 0
        dr("Clerk") = ""
        dr("Control") = ""
        dr("accountant") = ""
        dr("Cashier") = ""
        dr("CanBy") = "" 'm_CanBy
        dr("time1") = ""
        dr("Time2") = ""
        dr("Time3") = ""
        dr("CanTime") = ""
        dr("Interest") = 0
        dr("Receivable") = 0
        dr("Payable") = 0
        dr("Penal") = 0
        dr("Charges") = 0
        dr("Principle") = 0
        dr("TranIndicater") = ""
        dr("ComputerGenerated") = ""
        dr("SessionNumber") = ""
        dr("ReceiptNumber") = ""
        dr("SetNumber") = ""
        dr("CheckSum") = 0
        dr("MachineID") = ""
        dr("UserId") = "AUTO"
        dr("DateTimeCreation") = Now
        dr("CashTransactionID") = ""
        dr("CashTransactionID") = ""
        dr("Deleted") = "N"
        dr("Authenticated") = "1"
        dr("ChequeReasonReturnMasterID") = ""
        dr("Status") = ""
        dr("ClearingID") = ""
        dr("BrAdjEntryType") = ""
        dr("IsLog") = ""
        dr("PsStat") = ""
        dr("RecoDate") = DBNull.Value
        dr("RecoStatus") = ""

    End Sub

    Public Sub Save()

        Dim pk As ArrayList
        Dim dr As DataRow

        If ds.Tables.Count <= 0 Then
            Return
        End If

        pk = New ArrayList()
        pk.Add("DailyTransactionId")
        pk.Add("Date")
        BuildCommand(tbl.bnk_MainTransaction, ds.Tables(0), pk)

        For Each dr In ds.Tables(0).Rows
            If dr.RowState = DataRowState.Unchanged Then
            Else
                RemoveNull(dr)
            End If
        Next

        _objCon.BeginTrans(IsolationLevel.ReadCommitted)

        Try
            _objCon.SaveDataSet(ds, InsertCommand, DeleteCommand, UpdateCommand, SelectCommand, pmInsert, pmDelete, pmUpdate, tbl.bnk_MainTransaction)
        Catch ex As Exception
            _objCon.AbortTrans()
            Throw ex
        End Try

        _objCon.CommitTrans(False)

    End Sub

End Class

Public Class Transaction

    Public Property DailyTransactionId As String
    Public Property VoucherNo As String
    Public Property TrDate As Date
    Public Property BookCode As String

    Public Property GLCode As String
    Public Property AccountNumber As Long
    Public Property CustomerNo As Long
    Public Property BranchCode As String

    Public Property Amount As Double
    Public Property DrCr As String
    Public Property Narration As String

End Class
