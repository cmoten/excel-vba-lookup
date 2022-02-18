# Description
The attached excel workbook contains the example data. There is a named range for the dates in cells C4:F6

# VBA Code
Sub GetBuilding()

Dim currDate As Date
Dim dateRange As Range
Dim currCell As Range
Dim anchorCell As Range
Dim numDays As Long
Dim numColumns As Long
Dim numRows As Long
Dim building As String
Dim item As String

currDate = Date 'Todays date
Set dateRange = Range("Dates") 'named range in worksheet
Set anchorCell = Range("B3") 'to determine how many rows and columbns to offset

'This will loop through each row in the dates range by column
'So it will iterate from cell C4 to D4 to E4, etc. and then go to row 5
For Each currCell In dateRange
    numDays = currCell.Value - currDate
    'Notification logic test
    If numDays >= 0 And numDays <= 14 Then
        numColumns = currCell.Column - anchorCell.Column
        numRows = currCell.Row - anchorCell.Row
        building = currCell.Offset(0, -numColumns).Value
        item = currCell.Offset(-numRows, 0).Value
        MsgBox "Building: " & building & " " & "Item: " & item
    End If
Next currCell

End Sub
