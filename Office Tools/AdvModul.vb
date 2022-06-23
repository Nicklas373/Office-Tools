Module AdvModul
    Public Function RandomString(ByRef Length As String) As String
        Dim str As String = Nothing
        Dim rnd As New Random
        For i As Integer = 0 To Length
            Dim chrInt As Integer = 0
            Do
                chrInt = rnd.Next(30, 122)
                If (chrInt >= 48 And chrInt <= 57) Or (chrInt >= 65 And chrInt <= 90) Or (chrInt >= 97 And chrInt <= 122) Then
                    Exit Do
                End If
            Loop
            str &= Chr(chrInt)
        Next
        Return str
    End Function
    Public Function CompLevel(level As String) As String
        Dim compressionLevel As String
        If level = "No" Then
            compressionLevel = "-mx0"
            Return compressionLevel
        ElseIf level = "Low" Then
            compressionLevel = "-mx1"
            Return compressionLevel
        ElseIf level = "Fast" Then
            compressionLevel = "-mx3"
            Return compressionLevel
        ElseIf level = "Normal" Then
            compressionLevel = "-mx5"
            Return compressionLevel
        ElseIf level = "Maximum" Then
            compressionLevel = "-mx7"
            Return compressionLevel
        ElseIf level = "Ultra" Then
            compressionLevel = "-mx9"
            Return compressionLevel
        Else
            compressionLevel = "-mx5"
            Return compressionLevel
        End If
    End Function
    Public Function CompType(type As String) As String
        Dim compressionType As String
        If type = "7Z" Then
            compressionType = "-t7z"
            Return compressionType
        ElseIf type = "ZIP" Then
            compressionType = "-tzip"
            Return compressionType
        Else
            compressionType = "-tzip"
            Return compressionType
        End If
    End Function
    Public Function CompExt(type As String) As String
        Dim compressionExt As String
        If type = "7Z" Then
            compressionExt = "7z"
            Return compressionExt
        ElseIf type = "ZIP" Then
            compressionExt = "zip"
            Return compressionExt
        Else
            compressionExt = "zip"
            Return compressionExt
        End If
    End Function
    Public Function PassType(pass As String) As String
        Dim passwordType As String
        If pass = "Password (No Encrypt)" Then
            passwordType = "pass_noenc"
            Return passwordType
        ElseIf pass = "Password (SHA256)" Then
            passwordType = "pass_sha256"
            Return passwordType
        ElseIf pass = "Password (HMAC SHA256)" Then
            passwordType = "pass_hmac256"
            Return passwordType
        Else
            passwordType = "pass_enc"
            Return passwordType
        End If
    End Function
End Module
