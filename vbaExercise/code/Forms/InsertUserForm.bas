Private Sub UserForm_Initialize()
    'When the form is shown or cleared the data are cleaned and a new code is retrieved
    Dim lastRow As Long
    
    lastRow = Sheet2.getNextCode()
    CodeTextBox.Value = lastRow
    NameTextBox.Value = ""
    BirthTextBox.Value = ""
    EmailTextBox.Value = ""
    AddressTextBox.Value = ""
End Sub

Private Sub CancelButton_Click()
    'Unload the current form to close it
    Unload Me
End Sub

Private Sub ClearButton_Click()
    'Clear the form to be filled again with new data
    Call UserForm_Initialize
End Sub

Private Sub OKButton_Click()
    'When the user want to create a new user clicks this button
    Dim emptyRow As Long
    Dim ok As Integer
    
    ok = Sheet2.createUser(CodeTextBox.Value, NameTextBox.Value, BirthTextBox.Value, EmailTextBox.Value, AddressTextBox.Value)
    
    If ok = 0 Then
        Call UserForm_Initialize
    ElseIf ok = 2 Then
        BirthTextBox.Value = ""
    ElseIf ok = 3 Then
        EmailTextBox.Value = ""
    End If
End Sub



