Public Class CustomerDb

    Inherits SaveDB

    Private ds As DataSet
    Private dt As DataTable
    Private dr As DataRow

    Private tbl As New Tables

    Public Sub New(ByVal _conn As MyConnection)
        _objCon = _conn
    End Sub

    Public Sub Init(ByVal emptyDb As Boolean)

        Dim sql As String

        If emptyDb Then

            sql = ""
            sql = sql & " DELETE FROM " & tbl.bnkCustomerMaster & " "
            _objCon.ExecuteNonQuery(sql)

        End If

        sql = ""
        sql = sql & " SELECT * "
        sql = sql & " FROM " & tbl.bnkCustomerMaster & " "
        ds = _objCon.ExecDataSet(sql, CommandType.Text)

        ds.Tables(0).TableName = tbl.bnkCustomerMaster
        ds.Tables(0).PrimaryKey = New DataColumn() {ds.Tables(0).Columns("CustomerNo")}
        dt = ds.Tables(0)

    End Sub

    Public Function GetUpdatedDs() As DataSet
        Return ds
    End Function

    Public Sub Add(ByVal oCust As Customer)

        dr = dt.Select("CustomerNo=" & oCust.CustomerNo & " ").FirstOrDefault

        If IsNothing(dr) Then
            dr = dt.Rows.Add
        End If

        dr.Item("CustomerNo") = oCust.CustomerNo
        dr.Item("CustomerType") = oCust.CustomerType
        dr.Item("Ref") = oCust.Ref
        dr.Item("Name") = oCust.Name
        dr.Item("NameInLanguage") = ""
        dr.Item("AddressID") = oCust.AddressId
        dr.Item("TdsGlcode") = ""
        dr.Item("TdsAccountNo") = 0
        dr.Item("Address1") = ""
        dr.Item("Address2") = ""
        dr.Item("Address3") = ""
        dr.Item("PanNo") = ""
        dr.Item("MemberNo") = oCust.MemberNo
        dr.Item("InterestCategory") = ""
        dr.Item("TDSApplicable") = 0
        dr.Item("Reason") = ""
        dr.Item("BranchCode") = oCust.BranchCode
        dr.Item("Occupation") = oCust.Occupation
        dr.Item("DOB") = DBNull.Value
        If IsValidDate(oCust.DOB) Then
            dr.Item("DOB") = oCust.DOB
        End If
        dr.Item("DateOfOpening") = DBNull.Value
        If IsValidDate(oCust.DateOfOpening) Then
            dr.Item("DateOfOpening") = oCust.DateOfOpening
        End If
        dr.Item("EmployeeCode") = oCust.EmployeeCode
        dr.Item("PFNo") = ""
        dr.Item("DepartmentCode") = ""
        dr.Item("FathersName") = ""
        dr.Item("Designation") = ""
        dr.Item("Basic") = 0
        dr.Item("AccountStatus") = "A"
        dr.Item("ClosedDate") = DBNull.Value
        dr.Item("ClosedBy") = ""
        dr.Item("DateOfRetirement") = DBNull.Value
        dr.Item("SectionCode") = ""
        dr.Item("LastPostedRecoveryPeriod") = ""
        dr.Item("DateOfJoiningService") = DBNull.Value
        dr.Item("isAuthorised") = "1"
        dr.Item("AuthorisedBy") = ""
        dr.Item("AreaCode") = ""
        dr.Item("CasteCode") = DBNull.Value
        dr.Item("SubCasteCode") = ""
        dr.Item("AgeOnJoining") = 0
        dr.Item("BankCode") = ""
        dr.Item("BankGLCode") = ""
        dr.Item("BankAcNo") = 0
        dr.Item("Sex") = oCust.Sex
        dr.Item("ReasonForLeaving") = ""
        dr.Item("DateOfMeeting") = DBNull.Value
        dr.Item("MeetingResolutionNo") = 0
        dr.Item("YearsInService") = 0
        dr.Item("isExpired") = "N"
        'dr.Item("isLog") = "N"
        dr.Item("CustomerStatus") = "N"
        dr.Item("BloodGroup") = ""
        dr.Item("PayBillNo") = ""
        dr.Item("ICardNo") = ""
        dr.Item("DOA") = DBNull.Value
        dr.Item("IsMarried") = ""
        dr.Item("IsInsuranceOpted") = ""
        dr.Item("TelephoneNo") = ""
        dr.Item("WebPassword") = ""
        dr.Item("Email") = ""
        dr.Item("DividendGroup") = oCust.DividendGroup
        dr.Item("BookCode") = ""
        dr.Item("VoucherNo") = 0
        dr.Item("VoucherDate") = DBNull.Value
        dr.Item("EntryDate") = DBNull.Value
        dr.Item("SMSMobileNo") = oCust.SMSMobileNo
        dr.Item("LocCode") = ""
        dr.Item("ClubCode") = ""
        dr.Item("FamilyCustNo") = 0
        dr.Item("MothersName") = ""
        dr.Item("DirectorCode") = 0
        dr.Item("InsuranceDate") = DBNull.Value
        dr.Item("InsuranceNo") = ""
        dr.Item("HusbandsName") = ""

    End Sub

    Public Sub Save()

        Dim pk As ArrayList
        Dim dr As DataRow

        If ds.Tables.Count <= 0 Then
            Return
        End If

        pk = New ArrayList()
        pk.Add("CustomerNo")

        BuildCommand(tbl.bnkCustomerMaster, ds.Tables(0), pk)

        For Each dr In ds.Tables(0).Rows
            If dr.RowState = DataRowState.Unchanged Then
            Else
                RemoveNull(dr)
            End If
        Next

        _objCon.BeginTrans(IsolationLevel.ReadCommitted)

        Try
            _objCon.SaveDataSet(ds, InsertCommand, DeleteCommand, UpdateCommand, SelectCommand, pmInsert, pmDelete, pmUpdate, tbl.bnkCustomerMaster)
        Catch ex As Exception
            _objCon.AbortTrans()
            Throw ex
        End Try

        _objCon.CommitTrans(False)

    End Sub

End Class

Public Class Customer

    Public Property CustomerNo As Long
    Public Property CustomerType As String
    Public Property Ref As String
    Public Property Name As String
    Public Property NameInLanguage As String
    Public Property AddressId As String
    Public Property MemberNo As Long
    Public Property BranchCode As String
    Public Property Occupation As String
    Public Property DOB As Date
    Public Property DateOfOpening As Date
    Public Property EmployeeCode As String
    Public Property AccountStatus As String
    Public Property ClosedDate As Date
    Public Property ClosedBy As Date
    Public Property Sex As String
    'Public Property isLog") = "N"
    Public Property CustomerStatus As String
    Public Property SMSMobileNo As String
    Public Property DividendGroup As String

    Public Sub New()

        EmployeeCode = 0

    End Sub

End Class