Public Class AccountDb

    Inherits SaveDB

    Private ds As DataSet
    Private dt As DataTable
    Private dr As DataRow

    Private tbl As Tables

    Public Sub New(ByVal _conn As MyConnection)
        _objCon = _conn
    End Sub

    Public Sub Init(Optional ByVal emptyDb As Boolean = False)

        Dim sql As String

        If emptyDb Then

            sql = ""
            sql = sql & " DELETE FROM " & tbl.bnkAccountMaster & " "
            _objCon.ExecuteNonQuery(sql)

        End If

        sql = ""
        sql = sql & " SELECT * "
        sql = sql & " FROM " & tbl.bnkAccountMaster & " "
        'sql = sql & " WHERE 1=0 "
        ds = _objCon.ExecDataSet(sql, CommandType.Text)

        ds.Tables(0).TableName = tbl.bnkAccountMaster
        ds.Tables(0).PrimaryKey = New DataColumn() {ds.Tables(0).Columns("GLCODE"), ds.Tables(0).Columns("AccountNo")}
        dt = ds.Tables(0)

    End Sub

    Public Function GetUpdatedDs() As DataSet
        Return ds
    End Function

    Public Sub Add(ByVal oAccount As Account)

        dr = dt.Select("GLCode='" & oAccount.GlCode & "' AND AccountNo=" & oAccount.AccountNo & "").FirstOrDefault
        If IsNothing(dr) Then
            dr = dt.Rows.Add
        End If
        dr.Item("AccountMasterID") = ""
        dr.Item("GlCode") = oAccount.GLCode
        dr.Item("AccountNo") = oAccount.AccountNo
        dr.Item("Ref") = oAccount.Ref
        dr.Item("AccountName") = oAccount.AccountName
        dr.Item("NameinLanguage") = ""
        dr.Item("CustomerNo") = oAccount.CustomerNo
        dr.Item("CustomerType") = ""
        dr.Item("AddressID") = oAccount.AddressID
        dr.Item("LedgerNo") = ""
        dr.Item("LedgerFolioNo") = ""

        dr.Item("AccountOpeningDate") = DBNull.Value
        If IsValidDate(oAccount.AccountOpeningDate) Then
            dr.Item("AccountOpeningDate") = oAccount.AccountOpeningDate
        End If

        dr.Item("AccountClosingDate") = DBNull.Value
        dr.Item("AccountStatus") = "A"
        If IsValidDate(oAccount.AccountClosingDate) Then
            dr.Item("AccountClosingDate") = oAccount.AccountClosingDate
            dr.Item("AccountStatus") = "C"
        End If

        dr.Item("CompanyName") = ""
        dr.Item("MemberNo") = oAccount.MemberNo
        dr.Item("Occupation") = oAccount.Occupation
        dr.Item("AccountCategoryID") = "GENERAL"
        dr.Item("InterestCategoryID") = "GENERAL"
        dr.Item("AccountIndicaterID") = "OPERATIVE"
        dr.Item("UserID") = ""
        dr.Item("AuthenticateUserID") = ""
        dr.Item("OperatingInstructionID") = ""
        dr.Item("DateOfLastTransaction") = DBNull.Value
        dr.Item("AccountConstituentID") = ""
        dr.Item("TDSID") = ""
        dr.Item("MinorAccountDateOfBirth") = DBNull.Value
        dr.Item("ChequeBookIssued") = ""
        dr.Item("PassBookIssued") = ""
        dr.Item("InterestAccount") = 0
        dr.Item("InterestGLCode") = ""
        dr.Item("DateTimeCreation") = Date.Now
        dr.Item("Remark") = ""
        dr.Item("SpecialInstruction") = ""
        dr.Item("SexCode") = ""
        dr.Item("CasteCode") = ""
        dr.Item("AreaCode") = ""
        dr.Item("PassBookIssuedOn") = DBNull.Value
        dr.Item("InterestPostingFrequency") = ""
        dr.Item("OtherGlCode") = 1
        dr.Item("OtherAccountNo") = 0
        dr.Item("DOB") = DBNull.Value
        If oAccount.DOB <> Date.MinValue Then
            dr.Item("DOB") = oAccount.DOB
        End If
        dr.Item("CessationRemarks") = ""
        dr.Item("AgentCode") = 0
        'dr.Item("isLog") = "N"
        dr.Item("BranchCode") = oAccount.BranchCode
        dr.Item("Age") = 0
        dr.Item("SMSMobileNo") = ""

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
        BuildCommand(tbl.bnkAccountMaster, ds.Tables(0), pk)

        For Each dr In ds.Tables(0).Rows
            If dr.RowState = DataRowState.Unchanged Then
            Else
                RemoveNull(dr)
            End If
        Next

        _objCon.BeginTrans(IsolationLevel.ReadCommitted)

        Try
            _objCon.SaveDataSet(ds, InsertCommand, DeleteCommand, UpdateCommand, SelectCommand, pmInsert, pmDelete, pmUpdate, tbl.bnkAccountMaster)
        Catch ex As Exception
            _objCon.AbortTrans()
            Throw ex
        End Try

        _objCon.CommitTrans(False)

    End Sub

End Class

Public Class Account

    Public Property AccountMasterID As String
    Public Property GlCode As String
    Public Property AccountNo As Long
    Public Property Ref As String
    Public Property AccountName As String
    Public Property NameinLanguage As String
    Public Property CustomerNo As Long
    Public Property CustomerType As String
    Public Property AddressID As String
    Public Property LedgerNo As String
    Public Property LedgerFolioNo As String

    Public Property AccountOpeningDate As Date

    Public Property AccountClosingDate As Date
    Public Property AccountStatus As String

    Public Property CompanyName As String
    Public Property MemberNo As Long
    Public Property Occupation As String
    Public Property AccountCategoryID As String
    Public Property InterestCategoryID As String
    Public Property AccountIndicaterID As String
    Public Property UserID As String
    Public Property AuthenticateUserID As String
    Public Property OperatingInstructionID As String
    Public Property DateOfLastTransaction As Date
    Public Property AccountConstituentID As String
    Public Property TDSID As String
    Public Property MinorAccountDateOfBirth As Date
    Public Property ChequeBookIssued As String
    Public Property PassBookIssued As String
    Public Property InterestAccount As Long
    Public Property InterestGLCode As String
    Public Property DateTimeCreation As Date
    Public Property Remark As String
    Public Property SpecialInstruction As String
    Public Property SexCode As String
    Public Property CasteCode As String
    Public Property AreaCode As String
    Public Property PassBookIssuedOn As Date
    Public Property InterestPostingFrequency As String
    Public Property OtherGlCode As String
    Public Property OtherAccountNo As Long
    Public Property DOB As Date
    Public Property CessationRemarks As String
    Public Property AgentCode As String
    Public Property isLog As String
    Public Property BranchCode As String
    Public Property Age As Integer
    Public Property SMSMobileNo As String

    Public Sub New()


    End Sub

End Class
