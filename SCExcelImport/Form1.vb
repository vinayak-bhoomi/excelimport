Imports System.IO
Imports System.Reflection
Imports Microsoft.Win32

Public Class Form1

    Private m_Con As MyConnection
    Private m_BranchCode As String
    Private m_OpenDate As Date

    Private Const GLCODE As Single = 1
    Private Const ACCOUNT_NO As Single = 2
    Private Const ACCOUNT_NAME As Single = 3
    Private Const AMOUNT As Single = 4
    Private Const DR_CR As Single = 5
    Private Const NARRATION As Single = 6

    Private Sub InitPara()
        m_BranchCode = "01"
        m_OpenDate = CDate("2015-01-31")
    End Sub

    Public Enum DivColmn
        MemberNo = 1
        Year = 2
        Product = 3
        IntRate = 4
        Dividend = 5
        Status = 6
        MainGLCode = 7
        WarrantNo = 8

    End Enum

    Private Function CreateConnection() As Boolean

        m_Con = New MyConnection("MySQL", "sunil", "visioners", "root", "", "N")

        Try
            If m_Con.VerifyConnection() = False Then
                MessageBox.Show("Connectionstring Problem! ", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show("Connection Error occured! " & vbCrLf & ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
        Return True

    End Function

    Private Sub ImportBtn_Click(sender As System.Object, e As System.EventArgs) Handles ImportBtn.Click

        StatusLbl.Text = "Creating connection.."
        Me.Refresh()

        If CreateConnection() = False Then
            Return
        End If

        InitPara()

        Import_2()

    End Sub

    Private Sub Import_1()

        Dim glcodes As String
        Dim folderName As String
        Dim fileName As String

        Dim rowNum As Long
        Dim rowCount As Long
        Dim oCust As Customer
        Dim oCustDb As CustomerDb
        Dim oAcc As Account
        Dim oAccDb As AccountDb
        Dim oTrn As Transaction
        Dim oTrnDb As TransactionDb
        Dim oXls As ExcelFile

        Dim lCustNo As Long
        Dim lAddId As Long
        Dim sCustName As String
        Dim lAccNo As Long
        Dim lTrnId As Long

        Dim d1001Bal As Double
        Dim d2101Bal As Double
        Dim d2102Bal As Double
        Dim d6001Bal As Double
        Dim d6002Bal As Double

        glcodes = "1001"

        folderName = getAppDirectory()
        fileName = Path.Combine(folderName, "Data.xls")

        If Not File.Exists(fileName) Then
            MessageBox.Show(String.Format("File Error!{0} file not found", fileName), "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        StatusLbl.Text = "Opening excel file.."
        Me.Refresh()

        oXls = New ExcelFile(fileName)
        oCustDb = New CustomerDb(m_Con)
        oAccDb = New AccountDb(m_Con)
        oTrnDb = New TransactionDb(m_Con)

        oCustDb.Init(True)
        oAccDb.Init(True)
        oTrnDb.Init(True)

        rowCount = oXls.LastRow

        Do While True
            rowNum += 1

            If rowNum = rowCount Then
                Exit Do
            End If

            lCustNo = ToLng(oXls.Cells(rowNum, 1))
            sCustName = ToStr(oXls.Cells(rowNum, 4))
            lAddId = lCustNo
            lAccNo = lCustNo

            If lCustNo = 0 Then
                Continue Do
            End If

            d1001Bal = ToDbl(oXls.Cells(rowNum, 5))
            d2101Bal = ToDbl(oXls.Cells(rowNum, 6))
            d2102Bal = ToDbl(oXls.Cells(rowNum, 7))
            d6001Bal = ToDbl(oXls.Cells(rowNum, 8))
            d6002Bal = ToDbl(oXls.Cells(rowNum, 9))

            StatusLbl.Text = String.Format("Customers {0}/{1}", rowNum, rowCount)
            Me.Refresh()

            oCust = New Customer
            oCust.CustomerType = "M"
            oCust.CustomerNo = lCustNo
            oCust.Name = sCustName
            oCust.AddressId = lAddId
            oCust.MemberNo = lCustNo
            oCust.DateOfOpening = m_OpenDate
            oCust.BranchCode = m_BranchCode
            oCust.EmployeeCode = ToStr(oXls.Cells(rowNum, 2))
            oCust.DividendGroup = ToStr(oXls.Cells(rowNum, 3))
            oCustDb.Add(oCust)

            For Each glcode As String In glcodes.Split(",")

                StatusLbl.Text = String.Format("Creating account {0}/{1}", glcode, lAccNo)
                Me.Refresh()

                'Account
                oAcc = New Account
                oAcc.CustomerNo = lCustNo
                oAcc.GlCode = glcode
                oAcc.AccountNo = lAccNo
                oAcc.AccountName = sCustName
                oAcc.AddressID = lAddId
                oAcc.AccountOpeningDate = m_OpenDate
                oAcc.BranchCode = m_BranchCode
                oAccDb.Add(oAcc)

                'Opening Balance

                lTrnId += 1

                oTrn = New Transaction
                oTrn.DailyTransactionId = lTrnId
                oTrn.VoucherNo = 1
                oTrn.TrDate = m_OpenDate
                oTrn.BookCode = "TX"
                oTrn.GLCode = glcode
                oTrn.AccountNumber = lAccNo
                oTrn.BranchCode = m_BranchCode

                Select Case glcode
                    Case "1001"
                        oTrn.Amount = d1001Bal
                        oTrn.DrCr = "C"
                    Case "2101"
                        oTrn.Amount = d2101Bal
                        oTrn.DrCr = "C"
                    Case "2102"
                        oTrn.Amount = d2102Bal
                        oTrn.DrCr = "C"
                    Case "6001"
                        oTrn.Amount = d6001Bal
                        oTrn.DrCr = "D"
                    Case "6002"
                        oTrn.Amount = d6002Bal
                        oTrn.DrCr = "D"
                End Select

                oTrn.Narration = "BY OPENING BALANCE"
                oTrnDb.Add(oTrn)

            Next

        Loop

        StatusLbl.Text = "Saving data to database"
        Me.Refresh()

        oCustDb.Save()
        oAccDb.Save()
        oTrnDb.Save()

        oXls.Close()

        StatusLbl.Text = String.Format("Importing customers finished", rowNum, rowCount)
        Me.Refresh()
    End Sub

    Private Sub Import_2()

        Dim glcodes As String
        Dim folderName As String
        Dim fileName As String
        Dim sRef As String

        Dim rowNum As Long
        Dim rowCount As Long
        Dim oCust As Customer
        Dim oCustDb As CustomerDb
        Dim oAcc As Account
        Dim oAccDb As AccountDb
        Dim oTrn As Transaction
        Dim oTrnDb As TransactionDb
        Dim oAdd As Address
        Dim oAddDb As AddressDb
        Dim oXls As ExcelFile

        Dim lCustNo As Long
        Dim lAddId As Long
        Dim sCustName As String
        Dim lAccNo As Long
        Dim lTrnId As Long
        Dim sAddress As String

        Dim d1001Bal As Double
        Dim d2101Bal As Double
        Dim d2102Bal As Double
        Dim d6001Bal As Double
        Dim d6002Bal As Double

        glcodes = "1001"

        folderName = getAppDirectory()
        fileName = Path.Combine(folderName, "Data.xls")

        If Not File.Exists(fileName) Then
            MessageBox.Show(String.Format("File Error!{0} file not found", fileName), "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        StatusLbl.Text = "Opening excel file.."
        Me.Refresh()

        oXls = New ExcelFile(fileName)
        oCustDb = New CustomerDb(m_Con)
        oAccDb = New AccountDb(m_Con)
        oTrnDb = New TransactionDb(m_Con)
        oAddDb = New AddressDb(m_Con)

        oCustDb.Init(True)
        oAccDb.Init(True)
        oTrnDb.Init(True)
        oAddDb.Init(True)

        rowCount = oXls.LastRow

        Do While True
            rowNum += 1

            If rowNum = rowCount Then
                Exit Do
            End If

            lCustNo = ToLng(oXls.Cells(rowNum, 1))

            lAddId = rowNum
            lAccNo = lCustNo

            If lCustNo = 0 Then
                Continue Do
            End If

            sCustName = ToStr(oXls.Cells(rowNum, 2))
            sAddress = ToStr(oXls.Cells(rowNum, 3))
            m_BranchCode = Format(ToInt(oXls.Cells(rowNum, 4)), "00")
            d1001Bal = ToDbl(oXls.Cells(rowNum, 5))

            StatusLbl.Text = String.Format("Customers {0}/{1}", rowNum, rowCount)
            Me.Refresh()

            sRef = ""
            If sCustName.IndexOf("Miss") >= 0 Then
                sCustName = sCustName.Replace("Miss", "")
                sRef = "Miss"
            ElseIf sCustName.IndexOf("Mr.") = 0 Then
                sCustName = sCustName.Replace("Mr.", "")
                sRef = "Mr."
            ElseIf sCustName.IndexOf("Mrs.") = 0 Then
                sCustName = sCustName.Replace("Mrs.", "")
                sRef = "Mrs."
            ElseIf sCustName.IndexOf("Mrs,") = 0 Then
                sCustName = sCustName.Replace("Mrs,", "")
                sRef = "Mrs,"
            ElseIf sCustName.IndexOf("Nr.") = 0 Then
                sCustName = sCustName.Replace("Nr.", "")
                sRef = "Nr."
            ElseIf sCustName.IndexOf("Mrs") = 0 Then
                sCustName = sCustName.Replace("Mrs", "")
                sRef = "Mrs"
            ElseIf sCustName.IndexOf("Mr") = 0 Then
                sCustName = sCustName.Replace("Mr", "")
                sRef = "Mr"
            ElseIf sCustName.IndexOf("Miss") = 0 Then
                sCustName = sCustName.Replace("Miss", "")
                sRef = "Miss"
            End If


            oCust = New Customer
            oCust.CustomerType = "M"
            oCust.CustomerNo = lCustNo
            oCust.Ref = sRef
            oCust.Name = Trim(sCustName)
            oCust.AddressId = lAddId
            oCust.MemberNo = lCustNo
            oCust.DateOfOpening = m_OpenDate
            oCust.BranchCode = m_BranchCode
            oCustDb.Add(oCust)

            oAdd = New Address
            oAddDb.Add(lAddId, sAddress)

            For Each glcode As String In glcodes.Split(",")

                StatusLbl.Text = String.Format("Creating account {0}/{1}", glcode, lAccNo)
                Me.Refresh()

                'Account
                oAcc = New Account
                oAcc.CustomerNo = lCustNo
                oAcc.GlCode = glcode
                oAcc.AccountNo = lAccNo
                oAcc.AccountName = sCustName
                oAcc.AddressID = lAddId
                oAcc.AccountOpeningDate = m_OpenDate
                oAcc.BranchCode = m_BranchCode
                oAccDb.Add(oAcc)

                'Opening Balance

                lTrnId += 1

                oTrn = New Transaction
                oTrn.DailyTransactionId = lTrnId
                oTrn.VoucherNo = 1
                oTrn.TrDate = m_OpenDate
                oTrn.BookCode = "TX"
                oTrn.GLCode = glcode
                oTrn.AccountNumber = lAccNo
                oTrn.BranchCode = m_BranchCode

                Select Case glcode
                    Case "1001"
                        oTrn.Amount = d1001Bal
                        oTrn.DrCr = "C"
                    Case "2101"
                        oTrn.Amount = d2101Bal
                        oTrn.DrCr = "C"
                    Case "2102"
                        oTrn.Amount = d2102Bal
                        oTrn.DrCr = "C"
                    Case "6001"
                        oTrn.Amount = d6001Bal
                        oTrn.DrCr = "D"
                    Case "6002"
                        oTrn.Amount = d6002Bal
                        oTrn.DrCr = "D"
                End Select

                oTrn.Narration = "BY OPENING BALANCE"
                oTrnDb.Add(oTrn)

            Next

        Loop

        StatusLbl.Text = "Saving data to database"
        Me.Refresh()

        oCustDb.Save()
        oAccDb.Save()
        oTrnDb.Save()
        oAddDb.Save()


        oXls.Close()

        StatusLbl.Text = String.Format("Importing customers finished", rowNum, rowCount)
        Me.Refresh()
    End Sub

    Private Function getAppDirectory() As String
        Dim reslt As String

        reslt = Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName().CodeBase)
        Return Mid(reslt, 7)
    End Function

    Private Sub Form1_Disposed(sender As Object, e As System.EventArgs) Handles Me.Disposed

    End Sub


    Private Sub TestBtn_Click(sender As System.Object, e As System.EventArgs)

        Dim commaDel As String = "Belloi, Nuvem, P.O. Salcette - Goa., "
        Dim arryStr As List(Of String)

        arryStr = commaDel.Split(",").ToList

        Dim str As String

        For i As Integer = 0 To arryStr.Count - 1
            str = Trim(arryStr(i))
            If str = "" Then
                arryStr.RemoveAt(i)
            Else
                arryStr(i) = str
            End If
        Next

    End Sub

    Private Sub TranImportBtn_Click(sender As System.Object, e As System.EventArgs) Handles TranImportBtn.Click

        Dim rowNum As Long
        Dim xlsFileApp As ExcelFile
        Dim xlsTemptName As String
        Dim saveDlgBox As New OpenFileDialog
        Dim oTrn As Transaction_V1.Voucher
        Dim oTrnDb As Transaction_V1
        Dim totalRows As Integer
        Dim hasError As Boolean

        saveDlgBox.Filter = "Excel 97-2003 Workbook (*.xls)|*.xls|Excel Workbook (*.xlsx)|*.xlsx"
        saveDlgBox.RestoreDirectory = True

        If saveDlgBox.ShowDialog <> Windows.Forms.DialogResult.OK Then
            Return
        End If

        StatusLbl.Text = "Opening File..."

        xlsTemptName = saveDlgBox.FileName
        xlsFileApp = New ExcelFile(xlsTemptName)

        totalRows = xlsFileApp.LastRow

        hasError = False
        StatusLbl.Text = "Veryfing File..."

        rowNum = 1
        Do While True
            rowNum += 1
            If rowNum > totalRows Then
                Exit Do
            End If

            StatusLbl.Text = "Veryfing " & rowNum & "/" & totalRows
            StatusLbl.Refresh()

            If UtilityFunc.ToStr(xlsFileApp.Cells(rowNum, GLCODE)) = "" Then
                Continue Do
            End If

            If UtilityFunc.ToLng(xlsFileApp.Cells(rowNum, ACCOUNT_NO)) = 0 Then
                hasError = True
                MessageBox.Show("Error invalid account number!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Do
            End If

            If Not HasMatchAny(UtilityFunc.ToStr(xlsFileApp.Cells(rowNum, DR_CR)), "D", "C") Then
                hasError = True
                MessageBox.Show("Error invalid transaction type (Dr/Cr)!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Do
            End If
        Loop

        If hasError Then
            StatusLbl.Text = "Veryfication faild..."
            xlsFileApp.Close()
            Exit Sub
        End If

        StatusLbl.Text = "Veryfication done successfully..."

        '---------------------------------------------------------------

        StatusLbl.Text = "Import in proccess..."

        If Not GetConnectionByDSN("MySQL") Then
            Exit Sub
        End If

        Dim voucherDate As Date
        Dim voucherNo As String

        voucherDate = VoucherDateField.Value
        oTrnDb = New Transaction_V1(m_Con, voucherDate)
        voucherNo = oTrnDb.getNewVoucherNo("01", voucherDate)

        rowNum = 1
        Do While True
            rowNum += 1
            If rowNum > totalRows Then
                Exit Do
            End If

            If UtilityFunc.ToStr(xlsFileApp.Cells(rowNum, GLCODE)) = "" Then
                Continue Do
            End If

            StatusLbl.Text = "Import in proccess " & rowNum & "/" & totalRows
            StatusLbl.Refresh()

            oTrn = New Transaction_V1.Voucher
            oTrn.TrDate = voucherDate
            oTrn.BookCode = "TX"
            oTrn.VoucherNo = voucherNo
            oTrn.GLCode = xlsFileApp.Cells(rowNum, GLCODE)
            oTrn.AccountNo = xlsFileApp.Cells(rowNum, ACCOUNT_NO)
            oTrn.Amount = xlsFileApp.Cells(rowNum, AMOUNT)
            oTrn.DbCr = xlsFileApp.Cells(rowNum, DR_CR)
            oTrn.Narration = xlsFileApp.Cells(rowNum, NARRATION)
            oTrn.BranchCode = "01"
            oTrnDb.Add(oTrn)

        Loop

        oTrnDb.Save()
        xlsFileApp.Close()

        StatusLbl.Text = "Import done...."
    End Sub

    Private Sub TransTempltBtn_Click(sender As System.Object, e As System.EventArgs) Handles TransTempltBtn.Click

        Dim xlsFileApp As ExcelFile
        Dim xlsTemptName As String
        Dim saveDlgBox As New SaveFileDialog

        saveDlgBox.Filter = "Excel Workbook (*.xlsx)|*.xlsx|Excel 97-2003 Workbook (*.xls)|*.xls"
        saveDlgBox.RestoreDirectory = True

        If saveDlgBox.ShowDialog <> Windows.Forms.DialogResult.OK Then
            Return
        End If

        StatusLbl.Text = "Please wait..."

        xlsTemptName = saveDlgBox.FileName
        xlsFileApp = New ExcelFile(xlsTemptName)

        xlsFileApp.Cells(1, GLCODE) = "GLCode"
        xlsFileApp.Cells(1, ACCOUNT_NO) = "Account Number"
        xlsFileApp.Cells(1, ACCOUNT_NAME) = "Name (Optional)"
        xlsFileApp.Cells(1, AMOUNT) = "Amount"
        xlsFileApp.Cells(1, DR_CR) = "Dr/Cr"
        xlsFileApp.Cells(1, NARRATION) = "Narration"

        xlsFileApp.Save()
        xlsFileApp.Close()

        StatusLbl.Text = "Template is ready to use.."

        Process.Start(xlsTemptName)

    End Sub

    Private Sub DividendTempltBtn_Click(sender As System.Object, e As System.EventArgs) Handles DividendTempltBtn.Click

        Dim xlsFileApp As ExcelFile
        Dim xlsTemptName As String
        Dim saveDlgBox As New SaveFileDialog

        saveDlgBox.Filter = "Excel 97-2003 Workbook (*.xls)|*.xls|Excel Workbook (*.xlsx)|*.xlsx"
        saveDlgBox.RestoreDirectory = True

        If saveDlgBox.ShowDialog <> Windows.Forms.DialogResult.OK Then
            Return
        End If

        StatusLbl.Text = "Please wait..."

        xlsTemptName = saveDlgBox.FileName
        xlsFileApp = New ExcelFile(xlsTemptName)

        xlsFileApp.Cells(1, DivColmn.MemberNo) = "Member No"
        xlsFileApp.Cells(1, DivColmn.Year) = "Year"
        xlsFileApp.Cells(1, DivColmn.Product) = "Products"
        xlsFileApp.Cells(1, DivColmn.IntRate) = "Interest Rate"
        xlsFileApp.Cells(1, DivColmn.Dividend) = "Dividend"
        xlsFileApp.Cells(1, DivColmn.Status) = "Status"
        xlsFileApp.Cells(1, DivColmn.MainGLCode) = "Main GLCode"
        xlsFileApp.Cells(1, DivColmn.WarrantNo) = "Warrant No"

        xlsFileApp.Save()
        xlsFileApp.Close()

        StatusLbl.Text = "Template is ready to use.."

        Process.Start(xlsTemptName)

    End Sub

    Private Sub ImportDividendBtn_Click(sender As System.Object, e As System.EventArgs) Handles ImportDividendBtn.Click

        Dim rowNum As Long
        Dim xlsFApp As ExcelFile
        Dim xlsTemptName As String
        Dim saveDlgBox As New OpenFileDialog
        Dim oDiv As Dividend
        Dim oDivDb As DividendDb
        Dim totalRows As Integer
        Dim hasError As Boolean

        saveDlgBox.Filter = "Excel 97-2003 Workbook (*.xls)|*.xls|Excel Workbook (*.xlsx)|*.xlsx"
        saveDlgBox.RestoreDirectory = True

        If saveDlgBox.ShowDialog <> Windows.Forms.DialogResult.OK Then
            Return
        End If

        StatusLbl.Text = "Opening File..."

        xlsTemptName = saveDlgBox.FileName
        xlsFApp = New ExcelFile(xlsTemptName)

        totalRows = xlsFApp.LastRow

        hasError = False
        StatusLbl.Text = "Veryfing File..."

        rowNum = 1
        Do While True
            rowNum += 1
            If rowNum > totalRows Then
                Exit Do
            End If

            StatusLbl.Text = "Veryfing " & rowNum & "/" & totalRows
            StatusLbl.Refresh()

            If UtilityFunc.ToLng(xlsFApp.Cells(rowNum, DivColmn.MemberNo)) = 0 Then
                hasError = True
                MessageBox.Show("Error invalid member numner!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Do
            End If

            If UtilityFunc.ToStr(xlsFApp.Cells(rowNum, DivColmn.Year)) = "" Then
                hasError = True
                MessageBox.Show("Error invalid dividend Year!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Do
            End If

            If Not HasMatchAny(UtilityFunc.ToStr(xlsFApp.Cells(rowNum, DivColmn.Status)), "A", "P") Then
                hasError = True
                MessageBox.Show("Error invalid dividend status (A/P)!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Do
            End If
        Loop

        If hasError Then
            StatusLbl.Text = "Veryfication faild..."
            xlsFApp.Close()
            Exit Sub
        End If

        StatusLbl.Text = "Veryfication done successfully..."

        '---------------------------------------------------------------

        StatusLbl.Text = "Import in proccess..."

        If Not GetConnectionByDSN("MySQL") Then
            Exit Sub
        End If

        oDivDb = New DividendDb(m_Con)

        rowNum = 1
        Do While True
            rowNum += 1
            If rowNum > totalRows Then
                Exit Do
            End If

            StatusLbl.Text = "Import in proccess " & rowNum & "/" & totalRows
            StatusLbl.Refresh()

            oDiv = New Dividend
            oDiv.MemberNo = xlsFApp.Cells(rowNum, DivColmn.MemberNo)
            oDiv.Year = xlsFApp.Cells(rowNum, DivColmn.Year)
            oDiv.Product = xlsFApp.Cells(rowNum, DivColmn.Product)
            oDiv.IntRate = xlsFApp.Cells(rowNum, DivColmn.IntRate)
            oDiv.Dividend = xlsFApp.Cells(rowNum, DivColmn.Dividend)
            oDiv.Status = xlsFApp.Cells(rowNum, DivColmn.Status)
            oDiv.MainGlCode = UtilityFunc.ToStr(xlsFApp.Cells(rowNum, DivColmn.MainGLCode))
            oDiv.WarrantNo = UtilityFunc.ToInt(xlsFApp.Cells(rowNum, DivColmn.WarrantNo))
            oDivDb.Add(oDiv)

        Loop

        oDivDb.Save()
        xlsFApp.Close()

        StatusLbl.Text = "Import done...."

    End Sub

    Private Function HasMatchAny(ByVal objVal As Object, ParamArray compareList() As Object)

        If IsDBNull(objVal) Then
            Return False
        End If

        Dim match As Boolean

        For Each compValue As Object In compareList
            If IsDBNull(compValue) Then
                Continue For
            End If
            match = System.Text.RegularExpressions.Regex.IsMatch(objVal.ToString(), compValue.ToString(), System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            If match Then
                Return True
            End If
        Next

        Return False
    End Function

    Private Function GetConnectionByDSN(ByVal dsnName As String) As Boolean

        Dim DbType As String
        Dim ServerName As String
        Dim DbName As String
        Dim UserName As String
        Dim Password As String
        Dim Trust As String
        Dim regVersion As RegistryKey
        Dim hasError As Boolean

        'msgStr = "Connecting to branch"
        'RaiseEvent OnProgress(msgStr)

        regVersion = Registry.CurrentUser.OpenSubKey("Software\\ODBC\\ODBC.INI\\ODBC Data Sources", False)

        If (regVersion Is Nothing) Then
            'DSN Not Found
            MessageBox.Show("Connection Error! DSN not found " & dsnName, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

        'Check DBType
        If regVersion.GetValue(dsnName) = "" Then
            MessageBox.Show("Connection Error! DSN not found " & dsnName, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
            Return False
        End If

        If regVersion.GetValue(dsnName).ToString().Contains("MySQL") Then
            DbType = "MySQL"
        ElseIf regVersion.GetValue(dsnName).ToString().Contains("SQL") Then
            DbType = "SQL"
            Password = InputBox("Please enter SQL Server password", "SQL Password", "sa")
        Else
            'Invalid DB Type
            'msgStr = "Invalid Database Driver " & dsnName
            'RaiseEvent OnProgress(msgStr)
            Return False
        End If

        'Get DbName and details
        regVersion = Registry.CurrentUser.OpenSubKey("Software\\ODBC\\ODBC.INI\\" & dsnName, False)
        If regVersion Is Nothing Then
            Return False
        End If

        DbName = regVersion.GetValue("Database")
        ServerName = regVersion.GetValue("Server")
        UserName = regVersion.GetValue(IIf(DbType = "MySQL", "UID", "LastUser"))
        Trust = IIf(regVersion.GetValue("Trusted_Connection") = "Yes", "Y", "N")

        Try
            m_Con = New MyConnection(DbType, ServerName, DbName, UserName, Password, Trust)
            m_Con.VerifyConnection()
        Catch ex As Exception
            hasError = True
            MessageBox.Show("Error occured!", ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try

        Return True
    End Function

End Class

