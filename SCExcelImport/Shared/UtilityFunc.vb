Public Module UtilityFunc
    Public Function JoinStr(ByVal seprtChar As String, ParamArray strList As Object()) As String
        Dim result As String = String.Empty

        For Each strVal As Object In strList

            If IsNullOrEmpty(strVal) Then
                Continue For
            End If
            result += strVal
            If Not seprtChar = "" Then
                result += seprtChar
            End If
        Next
        If result = "" Then
            Return Nothing
        Else
            result = result.Remove(result.Length - seprtChar.Length)
            Return result
        End If

    End Function

    Public Function IsValidDate(ByVal dtObj As Object)
        If IsNothing(dtObj) Then
            Return False
        End If

        If Not IsDate(dtObj) Then
            Return False
        End If

        If dtObj = Date.MinValue Then
            Return False
        End If

        If dtObj <= CDate("1899-12-30") Then
            Return False
        End If

        Return True
    End Function

    Public Function IsNullOrEmpty(ByVal obj As Object) As Boolean

        If IsDBNull(obj) Then
            Return True
        End If

        If Trim(obj) = "" Then
            Return True
        End If

        Return False

    End Function

    Public Function IsNullOrZero(ByVal obj As Object) As Boolean

        If IsDBNull(obj) Then
            Return True
        End If

        If Val(obj) = 0 Then
            Return True
        End If

        Return False

    End Function

    Public Function ToInt(ByVal obj As Object) As Integer

        If IsNothing(obj) Then
            Return 0
        End If

        If IsDBNull(obj) Then
            Return 0
        End If

        If obj.ToString() = "" Then
            Return 0
        End If

        If Not IsNumeric(obj) Then
            Return 0
        End If

        Return CLng(obj.ToString())

    End Function

    Public Function ToLng(ByVal obj As Object) As Long

        If IsNothing(obj) Then
            Return 0
        End If

        If IsDBNull(obj) Then
            Return 0
        End If

        If obj.ToString() = "" Then
            Return 0
        End If

        If Not IsNumeric(obj) Then
            Return 0
        End If

        Return CLng(obj.ToString())

    End Function

    Public Function ToStr(ByVal obj As Object) As String

        If IsNothing(obj) Then
            Return ""
        End If

        If IsDBNull(obj) Then
            Return ""
        End If

        Return Trim(obj.ToString())

    End Function

    Public Function ToDbl(ByVal obj As Object) As Double

        If IsNothing(obj) Then
            Return 0
        End If

        If IsDBNull(obj) Then
            Return 0
        End If

        If obj.ToString() = "" Then
            Return 0
        End If

        If Not IsNumeric(obj) Then
            Return 0
        End If

        Return CDbl(obj.ToString())

    End Function

    'Public Function IsValidDate(ByVal objDate As Object) As Boolean
    '    If IsDBNull(objDate) Then
    '        Return False
    '    End If

    '    If Not IsDate(objDate) Then
    '        Return False
    '    End If

    '    If objDate = Date.MinValue Then
    '        Return False
    '    End If

    '    Return True
    'End Function
End Module
