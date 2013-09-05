Private Function GetComboItem(cmbItems As ComboBox, txtItem As String) As Integer
    Dim intLoop As Integer

    'If no elements in cmbItems then exit since there
    'is nothing to search for
    GetComboIndex = -1
    
    If cmbItems.ListCount = 0 Then
        Exit Function
    End If
    
    'Else search for the string
    For intLoop = 0 To cmbItems.ListCount
        If Trim(cmbItems.List(intLoop)) = Trim(txtItem) Then
            GetComboIndex = intLoop 'Item found
            Exit Function
        'Else
            'GetComboIndex = -1 'Item not found
        End If
    Next intLoop
End Function
