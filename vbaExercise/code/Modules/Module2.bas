Public Sub writeData(data As Range)
    'Write the range of data into a new workbook, checks if data is not empty and save it
    '
    '+ data(Range): range of data to be written in the new workbook
    Dim i As Integer
    Dim j As Integer
    Dim Wb2 As Workbook
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
    
    Set Wb2 = Workbooks.Add(1)
    i = 2
    j = 1
    
    With Wb2.Sheets(1)
        .Cells(1, 1).Value = "Code"
        .Cells(1, 2).Value = "Name"
        .Cells(1, 3).Value = "Birth"
        .Cells(1, 4).Value = "Email"
        .Cells(1, 5).Value = "Home Address"
        If data Is Nothing Then
            MsgBox ("No data found, try again.")
            Wb2.Close
            Exit Sub
        End If
        
        For Each res In data
            res.Copy (.Cells(i, j))
            If j = 5 Then
                j = 1
                i = i + 1
            Else
                j = j + 1
            End If
        Next
        Wb2.SaveAs Filename:=getRelativePath()
        Wb2.Close
    End With
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With
End Sub
 
Private Function getRelativePath() As String
    'Get a path to save the workbook with the current folder and a name based of time to avoid collisions
    '
    '- Return(String): valid name to save the new workbook
    getRelativePath = ThisWorkbook.Path & Application.PathSeparator & Format(Now, "yyyymmddhhnnss") & ".xlsx"
End Function
