
Public Class Tables

    'MASTER TABLE
    Public bnkCustomerMaster As String = "bnkCustomerMaster"
    Public mstNominee As String = "mstNominee"
    Public bnkAccountMaster As String = "bnkAccountMaster"
    Public bnkGLMaster As String = "bnkGLMaster"
    Public GLMap As String = "GLMap"
    Public mstAgentMaster As String = "mstAgentMaster"

    'DEPOSIT
    Public bnkRecurringdeposit As String = "bnkrecurringdeposit"
    Public bnkDepositReceipt As String = "bnkDepositReceipt"
    Public bnkDepositInterest As String = "bnkdepositinterest"
    Public bnkDailyDeposit As String = "bnkdailydeposit"
    Public bnkPledgedReceipt As String = "bnkpledgedreceipt"
    Public trnidgenerator As String = "trnidgenerator"
    'LOAN
    Public bnkLoanLimit As String = "bnkLoanLimit"
    Public bnkLoanAdvances As String = "bnkloanadvances"
    Public bnkLoaninterest As String = "bnkloaninterest"
    Public bnkDisbursement As String = "bnkdisbursement"
    Public bnkaccountinfo As String = "bnkaccountinfo"
    Public bnkSecurityGoldMaster As String = "bnksecuritygoldmaster"
    Public bnkSecurityDepositmaster As String = "bnksecuritydepositmaster"
    Public bnksecurityvehiclemaster As String = "bnksecurityvehiclemaster"
    Public bnkInsuranceMaster As String = "bnkInsuranceMaster"

    Public bnkSystemControlFile = "bnk_systemcontrolfile"

    Public mstAddress As String = "mstAddress"
    Public sys_idgenerator As String = "sys_idgenerator"
    Public idgenerator As String = "idgenerator"

    Public bnk_MainTransaction As String = "bnk_MainTransaction"

    Public bnk_trbalancefile As String = "bnk_trbalancefile"

    'SHARES
    Public mstSharesMaster As String = "mstSharesMaster"
    Public trnDividend As String = "trnDividend"

    Public objCon As MyConnection

    Sub New()
        'blank constructor
    End Sub

    Sub New(ByVal p_conn As MyConnection)
        objCon = p_conn
    End Sub

    Public Function bnkTransaction(ByVal workingDt As Date) As String
        If Format(workingDt, "dd/MM/yyyy") = Format(GetWorkingDate(), "dd/MM/yyyy") Then
            Return "bnk_trdailytransaction "
        Else
            Return "bnk_mainTransaction "
        End If
    End Function

    Private Function GetWorkingDate() As Date

        Dim sql As String
        Dim ds As DataSet


        sql = ""
        sql = sql & " Select * "
        sql = sql & " From " & bnkSystemControlFile
        sql = sql & " Where       Date=" & " ( "
        sql = sql & "                      Select Max(Date) From " & bnkSystemControlFile & " "
        sql = sql & "                    ) "
        ds = objCon.ExecDataSet(sql, CommandType.Text)

        If ds.Tables(0).Rows.Count <= 0 Then
            Return Format(Date.Now, "dd/MM/yyyy")
        ElseIf ds.Tables(0).Rows.Count > 1 Then

        ElseIf ds.Tables(0).Rows.Count = 1 Then
            Return ds.Tables(0).Rows(0).Item("Date")
        End If

        Return Format(Date.Now, "dd/MM/yyyy")
    End Function

End Class
