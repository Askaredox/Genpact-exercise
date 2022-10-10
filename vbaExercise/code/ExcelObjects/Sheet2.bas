Public Function createUser(code As Integer, name As String, birth As String, email As String, address As String) As Integer
    'Insert a user with the given data. This function also checks if the data given is correct.
    'Checks if the name and birth date are not empty, checks birth and email are correct
    '
    '+ code(Integer): code of the user to be inserted
    '+ name(String): name of the user [Required]
    '+ birth(String): birth of the user [Required] [Must be formated to "mm/dd/yyyy"]
    '+ email(String): email of the user [Must be a correct email address i.e. "aa@aa.com"]
    '+ address(String): address of the user
    '
    '- Return(Integer):exit code of the create user status, 0=ok, 1=data is missing, 2=date not valid, 3=email not valid
    Dim emptyRow As Integer
    Dim ok As Integer
    
    ok = checkData(name, birth, email)
    If ok <> 0 Then
        createUser = ok
        Exit Function
    End If
    
    emptyRow = getNextCode() + 1
    Cells(emptyRow, 1).Value = code
    Cells(emptyRow, 2).Value = name
    Cells(emptyRow, 3).Value = strToDate(birth)
    Cells(emptyRow, 4).Value = email
    Cells(emptyRow, 5).Value = address
End Function

Public Function readUser(What As String) As Range
    'Search for a user that have the same code as What or contains What in the name
    '
    '+ What(String): the code or keyword to be searched in the user database
    '
    '- Return(Range): all the users retrieved by the search
    Dim searchRange As Range
    Dim searchResult As Range
    
    Set searchRange = getRange()
    For i = 2 To searchRange.Rows.Count
        If CStr(Cells(i, 1).Value) = What Or Cells(i, 2).Value Like "*" & What & "*" Then
            If readUser Is Nothing Then
                Set readUser = Range("A" & i & ":E" & i)
            Else
                Set readUser = Union(readUser, Range("A" & i & ":E" & i))
            End If
        End If
       
    Next
End Function

Public Function updateUser(code As Integer, name As String, birth As String, email As String, address As String) As Boolean
    'Update a user with the given data. This function also checks if the data given is correct.
    'Checks if the name and birth date are not empty, checks birth and email are correct
    '
    '+ code(Integer): code of the user to be inserted
    '+ name(String): name of the user [Required]
    '+ birth(String): birth of the user [Required] [Must be formated to "mm/dd/yyyy"]
    '+ email(String): email of the user [Must be a correct email address i.e. "aa@aa.com"]
    '+ address(String): address of the user
    '
    '- Return(Boolean): return True=user edited or False=user not edited
    Dim lastRow As Long
    
    If checkData(name, birth, email) <> 0 Then
        editUser = False
        Exit Function
    End If
    
    lastRow = getNextCode()
    For i = 2 To lastRow
        If Cells(i, 1).Value = code Then
            Range("B" & i & ":E" & i).ClearContents
            Cells(i, 2).Value = name
            Cells(i, 3).Value = birth
            Cells(i, 4).Value = email
            Cells(i, 5).Value = address
            editUser = True
            Exit Function
        End If
    Next
    editUser = False
End Function

Public Function deleteUser(code As Integer) As Boolean
    'Delete a user with the given code
    '
    '+ code(Integer): code of the user to be deleted
    '
    '- Return(Boolean): return True=user deleted or False=user not deleted
    Dim lastRow As Long
    
    lastRow = getNextCode()
    For i = 2 To lastRow
        If Cells(i, 1).Value = code Then
            Range("A" & i & ":E" & i).ClearContents
            deleteUser = True
            Exit Function
        End If
    Next
    deleteUser = False
    
End Function


Public Function getFromTo(fromCode As Integer, toCode As Integer) As Range
    'Get all data from the users within the from and to code
    '
    '+ fromCode(Integer): from which code the list must be retrieved
    '+ toCode(Integer): to which code the list must be retrieved
    '
    '- Return(Range): all the users retrieved by the search
    Dim lastRow As Long
    Dim code As Integer
    
    lastRow = getNextCode()
    
    For i = 2 To lastRow
        code = Cells(i, 1).Value
        If fromCode <= code And code <= toCode Then
            If getFromTo Is Nothing Then
                Set getFromTo = Range("A" & i & ":E" & i)
            Else
                Set getFromTo = Union(getFromTo, Range("A" & i & ":E" & i))
            End If
        ElseIf code > toCode Then
            Exit Function
        End If
    Next
End Function

Private Function checkData(name As String, birth As String, email As String) As Integer
    'Checks if the data given is correct. Checks if the name and birth date are not empty,
    'checks birth and email are correct
    '
    '+ name(String): name of the user [Required]
    '+ birth(String): birth of the user [Required] [Must be formated to "mm/dd/yyyy"]
    '+ email(String): email of the user [Must be a correct email address i.e. "aa@aa.com"]
    '
    '- Return(Integer):exit code of check data status, 0=ok, 1=data is missing, 2=date not valid, 3=email not valid
    
    If (name = "") Or (birth = "") Then
        MsgBox ("Some data is required")
        checkData = 1
    ElseIf Not checkDate(birth) Then
        MsgBox ("Birth date is not valid")
        checkData = 2
    ElseIf Not isValidEmail(email) Then
        MsgBox ("Email address is not valid")
        checkData = 3
    Else
        checkData = 0
    End If
End Function

Private Function checkDate(date_data As String) As Boolean
    'Checks if the date given is in the correct format of "mm/dd/yyyy"
    '
    'date_data(String): date to be validated
    '
    '- Return(Boolean): return True=date ok or False=date not ok
    Dim ArrInput() As String
    ArrInput = Split(date_data, "/")
    
    Dim validDate As Boolean
    validDate = False
    
    If UBound(ArrInput) = 2 Then
        If ArrInput(0) > 0 And ArrInput(0) <= 12 And _
            ArrInput(1) > 0 And ArrInput(1) <= 31 Then
        validDate = True
        End If
    End If
    
    checkDate = validDate
    
End Function

Private Function isValidEmail(sEmailAddress As String) As Boolean
    'Checks if the email given is in the correct format of RegEx "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"
    '
    'sEmailAddress(String): email to be validated
    '
    '- Return(Boolean): return True=email ok or False=email not ok
    If sEmailAddress = "" Then
        isValidEmail = True
        Exit Function
    End If
    Dim sEmailPattern As String
    Dim oRegEx As Object
    Dim bReturn As Boolean
    
    sEmailPattern = "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"
    
    Set oRegEx = CreateObject("VBScript.RegExp")
    oRegEx.Global = True
    oRegEx.IgnoreCase = True
    oRegEx.Pattern = sEmailPattern
    
    isValidEmail = oRegEx.Test(sEmailAddress)
End Function

Public Function getNextCode() As Integer
    'Get the next row to be inserted a user in the table
    '
    '- Return(Integer): get the next possible valid correlative of a new user
    getNextCode = Cells(Rows.Count, 1).End(xlUp).row
End Function

Private Function getRange() As Range
    'Get the range of the table by the code
    '
    '- Return(Range): range of cells with the codes inside
    Dim row As Long
    Dim rangeStr As String
    
    row = getNextCode()
    
    rangeStr = "A2:A" & (row + 1)
    Set getRange = Range(rangeStr)
End Function

Private Function strToDate(dateStr As String) As Date
    'Convert a string into a date
    '+ dateStr(String): string to be converted to date
    '
    '- Return(Date): date of the given string
    Dim ArrInput() As String
    
    ArrInput = Split(dateStr, "/")
    
    strToDate = DateSerial(ArrInput(2), ArrInput(0), ArrInput(1))
End Function
