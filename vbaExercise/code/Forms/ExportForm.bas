Private Sub CommandButton1_Click()
    'Button of the "export range" to get the data and export it into a new Workbook
    Dim fromCode As Integer
    Dim toCode As Integer
    Dim data As Range
    
    fromCode = ComboBoxFrom.Value
    toCode = ComboBoxTo.Value
    Set data = Sheet2.getFromTo(fromCode, toCode)
    
    If Not checkCombo() Then
        ComboBoxFrom.Value = 1
        ComboBoxTo.Value = 1
        Exit Sub
    End If
    
    Call Module2.writeData(data)
    
    Unload Me
End Sub

Private Function checkCombo() As Boolean
    'Checks if the comboboxes of the from and to code are ok
    '
    '- Return(Boolean): True=ok or False=not ok
    Dim fromCode As Integer
    Dim toCode As Integer
    
    fromCode = ComboBoxFrom.Value
    toCode = ComboBoxTo.Value
    checkCombo = True
    
    If fromCode > toCode Then
        MsgBox ("From is greater than To, try again")
        checkCombo = False
    End If
End Function

Private Sub UserForm_Initialize()
    'When the userform is created the comboboxes are filled with valid codes to be retrieved
    For i = 1 To Sheet2.getNextCode() - 1
        ComboBoxFrom.AddItem (i)
        ComboBoxTo.AddItem (i)
    Next
    ComboBoxFrom.Value = 1
    ComboBoxTo.Value = 1
    
End Sub

