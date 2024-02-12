Sub FindAndColorAllInstances()
    Dim ws As Worksheet
    Dim cell As Range
    Dim searchText As String
    Dim cellText As String
    Dim startPos As Long

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("SAT Scheme Integration") ' Change to your sheet name

    ' Set the text string you want to find and color
    searchText = "CB8002" ' Change "YourText" to the text you want to find and color

    ' Set the color you want to apply
    Dim colorToApply As Long
    colorToApply = RGB(255, 0, 0) ' Change RGB values as per your desired color

    ' Loop through each cell in the worksheet
    For Each cell In ws.UsedRange
        ' Check if the cell contains the specified text
        cellText = cell.Value
        startPos = 1

        Do While startPos > 0
            startPos = InStr(startPos, cellText, searchText, vbTextCompare)
            If startPos > 0 Then
                ' Change the font color of each instance of the found text
                cell.Characters(startPos, Len(searchText)).Font.Color = colorToApply
				cell.Characters(startPos, Len(searchText)).Font.Bold = True
                startPos = startPos + Len(searchText)
            End If
        Loop
    Next cell
End Sub
