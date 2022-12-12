Sub TrimTo30Characters()
    Dim cell As Range
    Dim str As String
    Dim char As String

    ' Prompt the user to choose the 30th character for all cells
    char = InputBox("Choose the 30th character for all selected cells:", "Trim To 30 Characters")

    For Each cell In Selection
        str = cell.Value

        ' Check if the string is longer than 30 characters
        If Len(str) > 30 Then
            ' Trim the string to 30 characters and replace the 30th character with the chosen character
            str = Left(str, 29) & char
        Else
            ' Add the chosen character as the 30th character
            str = Left(str, 29) & char & Right(str, Len(str) - 29)
        End If

        ' Set the trimmed string as the cell value
        cell.Value = str
    Next
End Sub

'Written by https://chat.openai.com/chat
'Prompt "Write an VBA script that let you choose the 30th character in an dialog box on all selected cells, and trim all selected cells to 30 characters"
