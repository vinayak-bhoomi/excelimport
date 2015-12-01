Public Class GLDb
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
            sql = sql & " DELETE FROM " & tbl.bnkGLMaster & " "
            objCon.ExecuteNonQuery(sql)

        End If

        sql = ""
        sql = sql & " SELECT * "
        sql = sql & " FROM " & tbl.bnkGLMaster & " "
        sql = sql & " WHERE 1=0 "
        ds = objCon.ExecDataSet(sql, CommandType.Text)
        ds.Tables(0).TableName = tbl.bnkGLMaster

        dt = ds.Tables(0)
    End Sub

    Public Function GetUpdatedDs() As DataSet
        Return ds
    End Function

    Public Sub Add(ByVal oGL As GL)

        dr = dt.NewRow
        dr.Item("GLCode") = oGL.GLCode
        dr.Item("GLName") = oGL.Name
        dr.Item("GLNameInLanguage") = ""
        dr.Item("MainGLCode") = oGL.GLCode
        dr.Item("AutoGenAccountNo") = 0
        dr.Item("LastAccountNo") = 0
        dr.Item("CashBookCode") = ""
        dr.Item("DateTimeCreation") = DateTime.Now
        dr.Item("UserID") = "SA"
        dr.Item("InterestAccrualFrequency") = ""
        dr.Item("PenalAccrualFrequency") = ""
        dr.Item("InterestCalculationOn") = ""
        dr.Item("InterestCalculationMethod") = ""
        dr.Item("RoundingOffFactor") = 0
        dr.Item("MinimumProduct") = 0
        dr.Item("MinimumInterest") = 0
        dr.Item("MinimumBalancePeriodFromDays") = 0
        dr.Item("CapitalizeInterest") = 0
        dr.Item("TDSApplicable") = "N"
        dr.Item("InterestPostToOtherAC") = 0
        dr.Item("InterestPaidGL") = ""
        dr.Item("InterestPaidAccount") = 0
        dr.Item("InterestReceivedGL") = ""
        dr.Item("InterestReceivedAccount") = 0
        dr.Item("InterestRBLGL") = ""
        dr.Item("InterestRBLAccount") = 0
        dr.Item("InterestPBLGL") = ""
        dr.Item("InterestPBLAccount") = 0
        dr.Item("MBNPControlAccount") = 0
        dr.Item("MBNPGLCode") = ""
        dr.Item("MaintainIndInterestPayable") = 0
        dr.Item("MaintainIndInterestReceivable") = ""
        dr.Item("AutoPostingMaturityInterestProvision") = 0
        dr.Item("PenalInterestACGL") = ""
        dr.Item("PenalInterestACAccount") = 0
        dr.Item("PenalInterestRate") = 0
        dr.Item("IncidentalChargesACGL") = ""
        dr.Item("IncidentalChargesACAccount") = 0
        dr.Item("ApportionedAllowed") = 0
        dr.Item("InterestReceivable") = 0
        dr.Item("InterestAccrued") = 0
        dr.Item("PenalInterest") = 0
        dr.Item("IncidentalCharges") = 0
        dr.Item("AllowDrToLoanAccount") = 1
        dr.Item("ConsiderRecoveryInRBLODCalcu") = ""
        dr.Item("MinimumInterestPeriod") = 0
        dr.Item("LoanType") = ""
        dr.Item("IsRecoCalulateInterestDisbLastonth") = "N"
        dr.Item("RecoHalfMonthInterestDay") = 0
        dr.Item("IsRecoCalculateofSanctionedThisMonth") = "N"
        dr.Item("IsDeductInterestOnlyInTheFirstReco") = "N"
        dr.Item("IsMemberNoSameAsAccountNo") = "N"
        dr.Item("DeductInterestOnlyInTheFirstReco") = "N"
        dr.Item("ExcessRateforMaturedLoans") = 0
        'dr.Item("isLog") = "Y"
        dr.Item("ShortName") = ""
        dr.Item("IsCustomerAc") = "N"
        dr.Item("LiabilityBLGroupCode") = ""
        dr.Item("LiabilityBLGroupOrder") = 0
        dr.Item("AssetBLGroupCode") = 0
        dr.Item("AssetBLGroupOrder") = 0
        dr.Item("IsAutoCreateSubLedgerOnCustomerCreation") = "N"
        dr.Item("DailyGLShortCode") = ""
        dr.Item("EnableCollectionViaDDMachine") = "N"
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
        BuildCommand(tbl.bnkGLMaster, ds.Tables(0), pk)

        For Each dr In ds.Tables(0).Rows
            If dr.RowState = DataRowState.Unchanged Then
            Else
                RemoveNull(dr)
            End If
        Next

        objCon.BeginTrans(IsolationLevel.ReadCommitted)

        Try
            objCon.SaveDataSet(ds, InsertCommand, DeleteCommand, UpdateCommand, SelectCommand, pmInsert, pmDelete, pmUpdate, tbl.bnkGLMaster)
        Catch ex As Exception
            objCon.AbortTrans()
            Throw ex
        End Try

        objCon.CommitTrans(False)

    End Sub

End Class

Public Class GL

    Public Property GLCode As String
    Public Property Name As String
End Class