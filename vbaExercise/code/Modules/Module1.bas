Public Sub NewRecord()
    'show the form to create a new user
    InsertUserForm.Show
End Sub

Public Sub Search()
    'action of the button to search a list of users with the search value of a textbox
    Dim searchValue As String
    Dim searchResult As Range
    
    searchValue = Sheet1.TextBox1.Value
    
    Set searchResult = Sheet2.readUser(searchValue)
    Call Sheet1.showUsers(searchResult)
    
End Sub

Public Sub EditCode(code As Integer, row As Integer)
    'When the "Edit" button is pressed, a message box is shown and prepare data to be edited
    '
    '+ code(Integer): Code of the user to be edited
    '+ row(Integer): row of the data to be retrieved to edit a user
    Dim name As String
    Dim birth As String
    Dim email As String
    Dim address As String
    Dim ok As Boolean
    Dim respuesta As Integer
    
    respuesta = MsgBox("Are you sure you want to edit this record?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete prompt")
    If respuesta = vbYes Then
        name = Sheet1.getData(row, 3)
        birth = Format(Sheet1.getData(row, 4), "mm/dd/yyyy")
        email = Sheet1.getData(row, 5)
        address = Sheet1.getData(row, 6)
        
        ok = Sheet2.updateUser(code, name, birth, email, address)
        Call Search
    End If
End Sub

Public Sub DeleteCode(code As Integer)
    'When the "Delete" button is pressed, a message box is shown and prepare to delete a user
    '
    '+ code(Integer): Code of the user to be deleted
    Dim respuesta As Integer
    
    respuesta = MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete prompt")
    If respuesta = vbYes Then
        Sheet2.deleteUser (code)
        Call Search
    End If
End Sub

Public Sub ExportRange()
    'Show the form to export a range of users
    ExportForm.Show
End Sub

Public Sub ExportKeyword()
    'exports the current data shown in the table to be exported into a new workbook
    Dim searchValue As String
    Dim searchResult As Range
    
    searchValue = Sheet1.TextBox1.Value
    
    Set searchResult = Sheet2.readUser(searchValue)
    
    Call Module2.writeData(searchResult)
End Sub

