Module UtilityFunctn

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

End Module
