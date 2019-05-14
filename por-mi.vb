Sub cleanData()

Dim startRange As Range, resultRange As Range, tmp As Variant, book As Object
Dim sellerCodes As Range, idRange As Range, id As String, eventType As String, resultIDs As Range, resultNames As Range

Set book = Application.Workbooks("POR-Rev2")

Set startRange = book.Worksheets(1).Range("A2:A75")
Set resultRange = book.Worksheets(1).Range("E2:E75")

Set idRange = Application.Workbooks("user extract").Worksheets(1).Range("A2:A116738")
Set sellerCodes = Application.Workbooks("user extract").Worksheets(1).Range("B2:B116738")

Set resultIDs = book.Worksheets(1).Range("F2:F75")
Set resultNames = book.Worksheets(1).Range("G2:G75")

For counter = 1 To startRange.Count

    ' Parse the initial string
    tmp = parseDataString(startRange(counter, 1).Value, "-")
    id = Trim(tmp(1))
    For Each cell In idRange
        If id = cell.Value Then
            resultNames(counter, 1) = cell.Offset(0, 3)
            resultIDs(counter, 1) = cell.Offset(0, 1)
        End If

    Next cell

    resultRange(counter, 1) = id
Next counter

Debug.Print "Finished!"
End Sub


Function parseDataString(inputData As String, delimiter As String) As String()
    Dim all As Variant, returnVal(3) As String

    all = Split(inputData, delimiter)

    returnVal(0) = Trim(all(0))
    returnVal(1) = Trim(all(1))
    returnVal(2) = Trim(all(2))

    parseDataString = returnVal

End Function
