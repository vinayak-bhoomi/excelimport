Public Class AccountInfoDb
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
            sql = sql & " DELETE FROM " & tbl.bnkaccountinfo & " "
            sql = sql & " WHERE GLCode IN ('" & glcodes & "')"
            objCon.ExecuteNonQuery(sql)

        End If

        sql = ""
        sql = sql & " SELECT * "
        sql = sql & " FROM " & tbl.bnkaccountinfo & " "
        sql = sql & " WHERE GLCode IN ('" & glcodes & "')"
        ds = objCon.ExecDataSet(sql, CommandType.Text)

        ds.Tables(0).TableName = tbl.bnkaccountinfo
        ds.Tables(0).PrimaryKey = New DataColumn() {ds.Tables(0).Columns("GLCODE"), ds.Tables(0).Columns("AccountNumber"),
                                                    ds.Tables(0).Columns("MemberNo"), ds.Tables(0).Columns("LoanSerialNo")}
        dt = ds.Tables(0)

    End Sub

    Public Function GetUpdatedDs() As DataSet
        Return ds
    End Function

    Public Sub Add(ByVal oAcInfo As AccountInfo)

        dr = dt.Select("GLCode='" & oAcInfo.Glcode & "' " & _
                        " AND AccountNo=" & oAcInfo.Accountno & " " &
                        " AND MemberNo=" & oAcInfo.Memberno & "").FirstOrDefault

        If IsNothing(dr) Then
            dr = dt.Rows.Add
        End If

        dr.Item("id") = ""
        dr.Item("type") = oAcInfo.Type
        dr.Item("ref") = oAcInfo.Ref
        dr.Item("name") = oAcInfo.Name
        dr.Item("addressid") = oAcInfo.Addressid
        dr.Item("memberno") = oAcInfo.Memberno
        dr.Item("CustomerType") = "M"
        dr.Item("glcode") = oAcInfo.Glcode
        dr.Item("accountno") = oAcInfo.Accountno
        dr.Item("LoanSerialNo") = oAcInfo.LoanSerialNo
        dr.Item("glcode1") = ""
        dr.Item("accountno1") = 0
        dr.Item("userid") = ""
        dr.Item("datetimecreation") = DateTime.Now
        dr.Item("NameinLanguage") = ""
        dr.Item("relation") = ""
        dr.Item("NomineeRegNo") = ""
        dr.Item("RegDate") = oAcInfo.RegDate
        dr.Item("isActive") = "Y"
        'dr.Item("isLog") = "Y"
        dr.Item("Age") = 0

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
        pk.Add("MemberNo")

        BuildCommand(tbl.bnkaccountinfo, ds.Tables(0), pk)

        For Each dr In ds.Tables(0).Rows
            If dr.RowState = DataRowState.Unchanged Then
            Else
                RemoveNull(dr)
            End If
        Next

        objCon.BeginTrans(IsolationLevel.ReadCommitted)

        Try
            objCon.SaveDataSet(ds, InsertCommand, DeleteCommand, UpdateCommand, SelectCommand, pmInsert, pmDelete, pmUpdate, tbl.bnkaccountinfo)
        Catch ex As Exception
            objCon.AbortTrans()
            Throw ex
        End Try

        objCon.CommitTrans(False)

    End Sub

End Class

Public Class AccountInfo

    Public Property Id As String
    Public Property Type As String
    Public Property Ref As String
    Public Property Name As String
    Public Property Addressid As String
    Public Property Memberno As Long
    Public Property CustomerType As String
    Public Property Glcode As String
    Public Property Accountno As Long
    Public Property LoanSerialNo As Integer
    Public Property Glcode1 As String
    Public Property Accountno1 As Long

    Public Property Relation As String
    Public Property NomineeRegNo As String
    Public Property RegDate As Date
    Public Property Age As Integer

End Class