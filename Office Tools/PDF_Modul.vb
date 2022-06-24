Module PDF_Modul
    Public Function ImgCompLvlVal(Cmbx As String) As Integer
        'Combobox2.text'
        Dim value As Integer
        If Cmbx = "Highest" Then
            value = 20
        ElseIf Cmbx = "High" Then
            value = 40
        ElseIf Cmbx = "Normal" Then
            value = 60
        ElseIf Cmbx = "Low" Then
            value = 80
        ElseIf Cmbx = "Lowest" Then
            value = 100
        Else
            value = 0
        End If

        Return value
    End Function
    Public Function PdfIncUpdVal(Cmbx As Boolean) As Boolean
        'CheckBox5.Checked'
        Dim value As Boolean
        If Cmbx Then
            value = True
        Else
            value = False
        End If

        Return value
    End Function
    Public Function PdfMtOptVal(Cmbx As Boolean) As Boolean
        'CheckBox4.Checked'
        Dim value As Boolean
        If Cmbx Then
            value = True
        Else
            value = False
        End If

        Return value
    End Function
    Public Function PdfFoOptVal(Cmbx As Boolean) As Boolean
        'CheckBox2.Checked'
        Dim value As Boolean
        If Cmbx Then
            value = True
        Else
            value = False
        End If

        Return value
    End Function
    Public Function PdfOpcOptVal(Cmbx As Boolean) As Boolean
        'CheckBox3.Checked'
        Dim value As Boolean
        If Cmbx Then
            value = True
        Else
            value = False
        End If

        Return value
    End Function
End Module