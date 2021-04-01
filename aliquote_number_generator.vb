Sub
'Option Explicit 'Alla variabler måste sättas med Dim eller ReDim

Dim Row As Integer
Dim Key As String
Dim output As String
Dim aliquoteValue As Integer
Dim aliquoteCell As String
Dim previousKeyValue As String
Dim currentKeyValue As String

'Turn off screenupdate and calculations during script runtime.
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'First row containing value
Row = 2
'Cell that has the first KEY value
Key = Range("C2")
'Cell where we want first number to be written
output = Range("B2")
'Set variable
aliquoteValue = 0
'Set variable for comparision
previousKeyValue = ""
currentKeyValue = Key
'Set active cell
    'ActiveSheet.Range("B2").Activate() 'not needed since script works as intended without it

'Run until Key column is empty
Do Until Key = ""
    If currentKeyValue = previousKeyValue Then
        'Set current row
        aliquoteCell = "B" & Row
        'write value to cell
        Range(aliquoteCell) = aliquoteValue
    Else
        aliquoteCell = "B" & Row
        'increase value
        aliquoteValue = aliquoteValue + 1
        'write value to cell
        Range(aliquoteCell) = aliquoteValue
    End If
    
    'update variable with current value
    previousKeyValue = currentKeyValue
    'select the next cell to be updated.
        'ActiveCell.Offset(1, 0).Select() 'not needed since script works as intended without it

    Row = Row + 1
    'Set next cell in key column
    Key = Range("C" + CStr(Row))
    'update variable from current cell
    currentKeyValue = Key
        
Loop

'Turn on screenupdate and calculations
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub
