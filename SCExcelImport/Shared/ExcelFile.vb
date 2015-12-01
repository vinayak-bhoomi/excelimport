
Public Class ExcelFile

    Dim objXlsApp As Object
    Dim objXlsWorkbook As Object
    Dim objXlsSheet As Object

    Private _filePath As String
    Private _isNewFile As Boolean

    Public Sub New(ByVal filePath As String)

        _filePath = filePath

        objXlsApp = CreateObject("Excel.Application")
        If IO.File.Exists(filePath) Then
            objXlsWorkbook = objXlsApp.Workbooks.Open(filePath)
            _isNewFile = False
        Else
            objXlsWorkbook = objXlsApp.Workbooks.Add
            _isNewFile = True
        End If
        objXlsSheet = objXlsWorkbook.WorkSheets(1)

    End Sub

    'http://www.rondebruin.nl/mac/mac020.htm
    'http://www.rondebruin.nl/win/s5/win001.htm

    Public Sub Save()
        Dim fileExt As String

        fileExt = IO.Path.GetExtension(_filePath)

        If Not _isNewFile Then
            objXlsApp.ActiveWorkbook.Save()
        Else
            objXlsApp.ActiveWorkbook.SaveAs(_filePath, FileFormat:=56)
        End If
    End Sub

    Public Sub SaveAs(ByVal fileName As String)
        objXlsApp.ActiveWorkbook.SaveAs(fileName, FileFormat:=56)
    End Sub

    Public Sub Close()

        objXlsWorkbook.Close()
        objXlsApp.Quit()

        ReleaseObject(objXlsSheet)
        ReleaseObject(objXlsWorkbook)
        ReleaseObject(objXlsApp)

    End Sub

    Public Property Cells(ByVal row As Long, ByVal col As Long) As Object
        Get
            Return objXlsSheet.Cells(row, col).Value
        End Get
        Set(ByVal value As Object)
            objXlsSheet.Cells(row, col) = value
        End Set
    End Property

    Public ReadOnly Property LastRow() As Long
        Get
            If Not IsNothing(objXlsSheet) Then
                Dim _lastRow As Long
                _lastRow = objXlsSheet.UsedRange.Rows.Count
                Return _lastRow
            End If
            Return 0
        End Get
    End Property

    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            Dim intRel As Integer = 0
            Do
                intRel = System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            Loop While intRel > 0
        Catch ex As Exception
            MsgBox("Error while closing file " & ex.ToString)
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
End Class
