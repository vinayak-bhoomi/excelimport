Public Class DividendDb
    Inherits SaveDB

    Private _dr As DataRow
    Private _dt As DataTable
    Private _tbl As New Tables

    Sub New(ByVal myConn As MyConnection)
        _objCon = myConn

        Init()
    End Sub

    Private Sub Init()

        Dim sql As String

        sql = ""
        sql = sql & " SELECT * "
        sql = sql & " FROM " & _tbl.trnDividend & " "

        _dt = _objCon.ExecDataSet(sql, CommandType.Text).Tables(0)

        _dt.TableName = _tbl.trnDividend
        _dt.PrimaryKey = New DataColumn() {_dt.Columns("MemberNo"), _dt.Columns("Year"), _dt.Columns("MainGLCode")}

    End Sub

    Public Sub Add(ByVal dividend As Dividend)

        _dr = _dt.NewRow
        _dr("MemberNo") = dividend.MemberNo
        _dr("Year") = dividend.Year
        _dr("Product") = dividend.Product
        _dr("IntRate") = dividend.IntRate
        _dr("Dividend") = dividend.Dividend
        _dr("Status") = dividend.Status
        If dividend.PaidOn <> Date.MinValue Then
            _dr("PaidOn") = dividend.PaidOn
        End If
        _dr("GLCode") = dividend.GLCode
        _dr("AccountNo") = dividend.AccountNo
        _dr("BookCode") = dividend.BookCode
        _dr("VoucherNo") = dividend.VoucherNo
        _dr("WarrantNo") = dividend.WarrantNo
        If dividend.EPDate <> Date.MinValue Then
            _dr("EPDate") = dividend.EPDate
        End If
        _dr("EPMode") = dividend.EPMode
        _dr("DestBankCode") = dividend.DestBankCode
        _dr("DestAccountType") = dividend.DestAccountType
        _dr("DestAccountNo") = dividend.DestAccountNo
        _dr("DestReferenceNo") = dividend.DestReferenceNo
        _dr("MainGlCode") = dividend.MainGlCode
        _dr("ChequeNo") = dividend.ChequeNo
        If dividend.ChequeDate <> Date.MinValue Then
            _dr("ChequeDate") = dividend.ChequeDate
        End If
        _dt.Rows.Add(_dr)

    End Sub

    Public Sub Save()

        Dim pk As ArrayList

        pk = New ArrayList()
        pk.Add("MemberNo")
        pk.Add("Year")
        pk.Add("MainGLCode")

        BuildCommand(_tbl.trnDividend, _dt, pk)

        'For Each dr In _dt.Rows
        '    If dr.RowState = DataRowState.Unchanged Then
        '    Else
        '        RemoveNull(dr)
        '    End If
        'Next

        _objCon.BeginTrans(IsolationLevel.ReadCommitted)

        Try
            _objCon.SaveDataSet(_dt.DataSet, InsertCommand, DeleteCommand, UpdateCommand, SelectCommand, pmInsert, pmDelete, pmUpdate, _tbl.trnDividend)
        Catch ex As Exception
            _objCon.AbortTrans()
            Throw ex
        End Try

        _objCon.CommitTrans(False)

    End Sub

End Class


Public Class Dividend

    Public Property MemberNo As Long
    Public Property Year As String
    Public Property Product As Double
    Public Property IntRate As Integer
    Public Property Dividend As Double
    Public Property Status As String
    Public Property PaidOn As Date
    Public Property GLCode As String
    Public Property AccountNo As Long
    Public Property BookCode As String
    Public Property VoucherNo As Integer
    Public Property WarrantNo As Integer
    Public Property EPDate As Date
    Public Property EPMode As String
    Public Property DestBankCode As String
    Public Property DestAccountType As String
    Public Property DestAccountNo As String
    Public Property DestReferenceNo As String
    Public Property MainGlCode As String
    Public Property ChequeNo As String
    Public Property ChequeDate As Date

    Sub New()

        MemberNo = 0
        Year = ""
        Product = 0
        IntRate = 0
        Dividend = 0
        Status = "A"
        PaidOn = Date.MinValue
        GLCode = ""
        AccountNo = 0
        BookCode = ""
        VoucherNo = 0
        WarrantNo = 0
        EPDate = Date.MinValue
        EPMode = ""
        DestBankCode = ""
        DestAccountType = ""
        DestAccountNo = ""
        DestReferenceNo = ""
        MainGlCode = ""
        ChequeNo = ""
        ChequeDate = Date.MinValue

    End Sub

End Class