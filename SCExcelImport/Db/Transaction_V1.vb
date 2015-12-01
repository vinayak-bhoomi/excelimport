Public Class Transaction_V1
    Inherits SaveDB

    Dim dsTrn As New DataSet
    Dim dtTrn As DataTable
    Dim _tables As Tables

    Private _voucherNo As Long
    Private _bookCode As String
    Private _trnDate As Date


    Public Sub New(ByVal _conn As MyConnection, ByVal trdate As Date)

        _tables = New Tables(_conn)

        _objCon = _conn
        _trnDate = trdate

        InitTransactionDs()
    End Sub

    Private Sub InitTransactionDs()

        Dim sql As String

        sql = ""
        sql = sql & " SELECT * "
        sql = sql & " FROM " & _tables.bnkTransaction(_trnDate)
        sql = sql & " WHERE Date = '" & Format(_trnDate, "yyyy-MM-dd") & "'"
        sql = sql & " AND 1=0 "
        dsTrn = _objCon.ExecDataSet(sql, CommandType.Text)

        dsTrn.Tables(0).TableName = _tables.bnkTransaction(_trnDate)
        dsTrn.Tables(0).PrimaryKey = New DataColumn() {dsTrn.Tables(0).Columns("DailyTransactionID"), dsTrn.Tables(0).Columns("Date")}

    End Sub

    Public Sub Add(ByVal voucher As Voucher)

        Dim dr As DataRow

        dr = dsTrn.Tables(0).NewRow
        dr.Item("DailyTransactionID") = getNewTrnId(_trnDate)
        dr.Item("VoucherNo") = voucher.VoucherNo
        dr.Item("BookCode") = voucher.BookCode
        dr.Item("Date") = Format(voucher.TrDate, "dd/MM/yyyy")
        dr.Item("SerialNumber") = 0
        dr.Item("GLCode") = voucher.GLCode
        dr.Item("AccountNumber") = voucher.AccountNo
        dr.Item("CustomerNo") = voucher.CustomerNo
        dr.Item("DbCr") = voucher.DbCr
        dr.Item("Amount") = voucher.Amount
        dr.Item("Narration") = voucher.Narration
        dr.Item("BookType") = "X"
        dr.Item("BankCode") = ""
        dr.Item("BranchCode") = voucher.BranchCode
        dr.Item("PresentDate") = DBNull.Value
        dr.Item("DueDate") = DBNull.Value
        dr.Item("InstrumentAmount") = 0
        dr.Item("NarrationinMarathi") = ""
        dr.Item("InstrumentNumber") = 0
        dr.Item("InstrumentDate") = DBNull.Value
        dr.Item("InstrumentType") = ""
        dr.Item("SlipNo") = 0
        dr.Item("Clerk") = ""
        dr.Item("Control") = ""
        dr.Item("Accountant") = ""
        dr.Item("Cashier") = ""
        dr.Item("CanBy") = ""
        dr.Item("Time1") = ""
        dr.Item("Time2") = ""
        dr.Item("Time3") = ""
        dr.Item("CanTime") = ""
        dr.Item("Interest") = 0
        dr.Item("Receivable") = 0
        dr.Item("Payable") = 0
        dr.Item("Penal") = 0
        dr.Item("Charges") = 0
        dr.Item("Principle") = 0
        dr.Item("TranIndicater") = ""
        dr.Item("ComputerGenerated") = ""
        dr.Item("SessionNumber") = 0
        dr.Item("TokenNumber") = 0
        dr.Item("ReceiptNumber") = 0
        dr.Item("SetNumber") = 0
        dr.Item("CheckSum") = 0
        dr.Item("MachineID") = ""
        dr.Item("UserID") = "AUTO"
        dr.Item("DateTimeCreation") = DateTime.Now
        dr.Item("CashTransactionID") = 0
        dr.Item("Deleted") = "N"
        dr.Item("Authenticated") = "1"
        dr.Item("ChequeReasonReturnMasterID") = 0
        dr.Item("Status") = ""
        dr.Item("ClearingID") = ""
        dr.Item("BrAdjEntryType") = ""
        dr.Item("isLog") = ""
        dr.Item("psStat") = ""
        dr.Item("recoStatus") = ""
        dr.Item("recoDate") = DBNull.Value
        dr.Item("VoucherType") = ""
        dsTrn.Tables(0).Rows.Add(dr)

    End Sub

    Public Function Save() As Boolean


        Dim pk As ArrayList
        Dim dr As DataRow


        pk = New ArrayList(0)
        pk.Add("DailyTransactionID")
        pk.Add("Date")

        BuildCommand(_tables.bnkTransaction(_trnDate), dsTrn.Tables(0), pk)

        _objCon.BeginTrans(System.Data.IsolationLevel.ReadCommitted)

        Try
            _objCon.SaveDataSet(dsTrn, InsertCommand, DeleteCommand, UpdateCommand, SelectCommand, pmInsert, pmDelete, pmUpdate, _tables.bnkTransaction(_trnDate))
        Catch ex As Exception
            _objCon.AbortTrans()
            MessageBox.Show("Error Occured!" & ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try

        _objCon.CommitTrans()
        dsTrn.AcceptChanges()

        Return True
    End Function

    Public Function getNewVoucherNo(ByVal branchCode As String, ByVal trnDate As Date)

        Dim sql As String
        Dim result As Object
        Dim AcYear As String = ""
        Dim voucherType As String
        Dim idType As String
        Dim month As Integer

        idType = "V"
        voucherType = "J"
        month = 0

        If trnDate.Month <= 3 Then
            AcYear = Trim(trnDate.Year - 1).Substring(0, 4) & Trim(trnDate.Year).Substring(2, 2)
        Else
            AcYear = Trim(trnDate.Year).Substring(0, 4) & Trim(trnDate.Year + 1).Substring(2, 2)
        End If

        _objCon.BeginTrans(IsolationLevel.Serializable)

        sql = ""
        sql = sql & " SELECT ID "
        sql = sql & " FROM " & _tables.sys_idgenerator
        sql = sql & " WHERE IDTag1='" & idType & "' "
        sql = sql & " AND   IDTag2='" & voucherType & "' "
        sql = sql & " AND   IDTag3='" & branchCode & "' "
        sql = sql & " AND   AcYear='" & AcYear & "' "
        sql = sql & " AND   Month=" & month & ""
        sql = sql & " AND   Date Is NULL "
        result = _objCon.ExecuteScalar(sql)
        result += 1

        sql = ""
        sql = sql & " UPDATE " & _tables.sys_idgenerator
        sql = sql & " SET ID=" & result & ""
        sql = sql & " WHERE IDTag1='" & idType & "' "
        sql = sql & " AND   IDTag2='" & voucherType & "' "
        sql = sql & " AND   IDTag3='" & branchCode & "' "
        sql = sql & " AND   AcYear='" & AcYear & "' "
        sql = sql & " AND   Month=" & month & ""
        sql = sql & " AND   Date Is NULL "
        _objCon.ExecuteNonQuery(sql)

        _objCon.CommitTrans(False)

        Return result

    End Function

    Private Function getNewTrnId(ByVal dt As Date)

        Dim sql As String
        Dim objId As Object

        sql = ""
        sql = sql & " SELECT Number "
        sql = sql & " FROM " & _tables.trnidgenerator
        sql = sql & " WHERE TableName = 'BNK_TRDAILYTRANSACTION' "
        sql = sql & " AND   Date='" & Format(dt, "yyyy-MM-dd") & "'  "
        objId = _objCon.ExecuteScalar(sql)

        If IsDBNull(objId) Or IsNothing(objId) Then
            sql = ""
            sql = sql & " INSERT INTO " & _tables.trnidgenerator & " "
            sql = sql & "       (TableName,Prefix,PrefixAllowed,Date,Number,Length) "
            sql = sql & " Values"
            sql = sql & "       ('BNK_TRDAILYTRANSACTION','','N','" & Format(dt, "yyyy-MM-dd") & "',1,7)"
            _objCon.ExecuteNonQuery(sql)
            objId = 1
        Else
            objId += 1
            sql = ""
            sql = sql & " UPDATE " & _tables.trnidgenerator
            sql = sql & " SET Number =" & objId & ""
            sql = sql & " WHERE TableName = 'BNK_TRDAILYTRANSACTION' "
            sql = sql & " AND   Date='" & Format(dt, "yyyy-MM-dd") & "'  "
            _objCon.ExecuteNonQuery(sql)
        End If

        Return Format(objId, "0000000")
    End Function

    Public Class Voucher

        Public Property TrDate As Date

        Public UserName
        Private _voucherNo As Long
        Public Property VoucherNo() As Long
            Get
                Return _voucherNo
            End Get
            Set(ByVal value As Long)
                _voucherNo = value
            End Set
        End Property

        Private _bookCode As String
        Public Property BookCode() As String
            Get
                Return _bookCode
            End Get
            Set(ByVal value As String)
                _bookCode = value
            End Set
        End Property

        Private _glcode As String
        Public Property GLCode() As String
            Get
                Return _glcode
            End Get
            Set(ByVal value As String)
                _glcode = value
            End Set
        End Property

        Private _accountNo As Long
        Public Property AccountNo() As Long
            Get
                Return _accountNo
            End Get
            Set(ByVal value As Long)
                _accountNo = value
            End Set
        End Property

        Private _trnDate As Date
        Public Property TrnDate() As Date
            Get
                Return _trnDate
            End Get
            Set(ByVal value As Date)
                _trnDate = value
            End Set
        End Property

        Private _amount As Double
        Public Property Amount() As Double
            Get
                Return _amount
            End Get
            Set(ByVal value As Double)
                _amount = value
            End Set
        End Property

        Private _dbCr As String
        Public Property DbCr() As String
            Get
                Return _dbCr
            End Get
            Set(ByVal value As String)
                _dbCr = value
            End Set
        End Property

        Private _customerNo As Long
        Public Property CustomerNo() As Long
            Get
                Return _customerNo
            End Get
            Set(ByVal value As Long)
                _customerNo = value
            End Set
        End Property


        Private _narration As String
        Public Property Narration() As String
            Get
                Return _narration
            End Get
            Set(ByVal value As String)
                _narration = value
            End Set
        End Property

        Private _bookType As String
        Public Property BookType() As String
            Get
                Return _bookType
            End Get
            Set(ByVal value As String)
                _bookType = value
            End Set
        End Property

        Private _branchCode As String
        Public Property BranchCode() As String
            Get
                Return _branchCode
            End Get
            Set(ByVal value As String)
                _branchCode = value
            End Set
        End Property

    End Class

End Class
