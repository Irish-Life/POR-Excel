Sub cleanData()

    Dim startRange As Range
    Dim idRange As Range
    Dim sellerCodes As Range

    Dim tmp As Variant

    Dim id As String
    Dim eventLabel As String
    Dim email As String
    Dim eventTotal As String

    Dim resultIDs As Range
    Dim resultNames As Range
    Dim resultSellerCodes As Range
    Dim resultEventLabels As Range
    Dim resultSegments As Range
    Dim resultEventTotals As Range

    Set startRange = Application.Workbooks("POR - Usage_Events_Table").Worksheets(1).Range("A2:A207")

    Set idRange = Application.Workbooks("user extract").Worksheets(1).Range("A2:A116738")
    Set sellerCodes = Application.Workbooks("user extract").Worksheets(1).Range("B2:B116738")

    Set resultIDs = Application.Workbooks("Point-of-Retirement-Reporting").Worksheets(1).Range("A2:A207")
    Set resultNames = Application.Workbooks("Point-of-Retirement-Reporting").Worksheets(1).Range("B2:B207")
    Set resultSellerCodes = Application.Workbooks("Point-of-Retirement-Reporting").Worksheets(1).Range("C2:C207")
    Set resultEventLabels = Application.Workbooks("Point-of-Retirement-Reporting").Worksheets(1).Range("D2:D207")
    Set resultSegments = Application.Workbooks("Point-of-Retirement-Reporting").Worksheets(1).Range("E2:E207")
    Set resultEventTotals = Application.Workbooks("Point-of-Retirement-Reporting").Worksheets(1).Range("E2:E207")

    For counter = 1 To startRange.Count

        ' Parse the initial string
        tmp = parseDataString(startRange(counter, 1).Value, "-")
        eventLabel = Trim(tmp(0))
        id = Trim(tmp(1))
        segment = Trim(tmp(2))
        eventTotal = startRange(counter, 1).Offset(0, 4)

        For Each cell In idRange
            If id = cell.Value Then
                email = idRange(counter, 1).Offset(0, 4)
                If InStr(email, "irishlife.ie") > 0 Then Exit For

                resultNames(counter, 1) = cell.Offset(0, 3)
                resultIDs(counter, 1) = cell.Offset(0, 1)
                resultSellerCodes(counter, 1) = id
                resultEventLabels(counter, 1) = eventLabel
                resultSegments(counter, 1) = segment
                resultEventTotals(counter, 1) = eventTotal
            End If

        Next cell

    Next counter

    Debug.Print "Finished!"
End Sub


Function parseDataString(inputData As String, delimiter As String) As String()
    Dim all As Variant, returnVal(2) As String

    all = Split(inputData, delimiter)

    returnVal(0) = Trim(all(0))
    returnVal(1) = Trim(all(1))
    returnVal(2) = Trim(all(2))

    parseDataString = returnVal

End Function


