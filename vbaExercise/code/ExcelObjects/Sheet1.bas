Public Sub showUsers(searchResult As Range)
    'Show the users that have been get as a search result
    '
    '+ searchResult(Range): the users to be shown in the sheet
    Dim code As Integer
    Dim i As Integer
    Dim j As Integer
    
    Call clearAll
        
    If searchResult Is Nothing Then
        Cells(7, 2).Value = "No Result found"
    Else
        i = 7
        j = 2
        For Each res In searchResult
            res.Copy (Cells(i, j))
            If j = 2 Then code = Cells(i, j)
            If j = 6 Then
                Call insertButtons(code, i, j + 1, "Edit")
                Call insertButtons(code, i, j + 2, "Delete")
                j = 2
                i = i + 1
            Else
                j = j + 1
            End If
        Next
    End If
End Sub

Private Sub clearAll()
    'Clear the table before filling again with data including the
    'created buttons
    Dim TotalRange As Range
    
    Set TotalRange = usedRange
    Set TotalRange = TotalRange.Offset(6, 0).Resize(TotalRange.Rows.Count - 6, _
                                           TotalRange.Columns.Count)
    TotalRange.ClearContents
    
    For Each btn In Buttons
        
        If btn.Text <> "Add new user" And btn.Text <> "Search" And _
            btn.Text <> "Export from search" And btn.Text <> "Export range" Then
            btn.Delete
        End If
    Next
End Sub

Private Sub insertButtons(code As Integer, i As Integer, j As Integer, bType As String)
    'insert the buttons to edit or delete the entry
    '
    '+ code(Integer): code of the user to be edited or deleted
    '+ i(Integer): row of the entry to get the position to create the button
    '+ j(Integer): column of the entry to get the position to create the button
    '+ bType(String): type of the button to be created ["Edit","Delete"]
    Dim t As Range
    Dim btn As Button
    
    Set t = Range(Cells(i, j), Cells(i, j))
    Set btn = Buttons.Add(t.Left, t.Top, t.Width, t.Height)
    With btn
        If bType = "Edit" Then
            .OnAction = "'EditCode " & code & "," & i & "'"
        Else
            .OnAction = "'DeleteCode " & code & "'"
        End If
        .Caption = bType
        .name = bType & code
    End With
        
End Sub

Public Function getData(row As Integer, col As Integer) As Range
    'Obtain the data from a user with the row and column
    '
    '+ row(Integer): row to be retrieved
    '+ col(Integer): column to be retrieved
    '
    '- Return(Range): data of the cell as range
    getData = Cells(row, col)
End Function

