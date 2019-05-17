Sub cleanData()
    ' Declare variables
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


    ' Initialise variables
    Set startRange = Application.Workbooks("POR - Usage_Events_Table").Worksheets(1).Range("A2:A316")

    Set idRange = Application.Workbooks("user extract").Worksheets(1).Range("A2:A116738")
    Set sellerCodes = Application.Workbooks("user extract").Worksheets(1).Range("B2:B116738")

    Set resultIDs = Application.Workbooks("Point-of-Retirement-Reporting").Worksheets(1).Range("A2:A316")
    Set resultNames = Application.Workbooks("Point-of-Retirement-Reporting").Worksheets(1).Range("B2:B316")
    Set resultSellerCodes = Application.Workbooks("Point-of-Retirement-Reporting").Worksheets(1).Range("C2:C316")
    Set resultEventLabels = Application.Workbooks("Point-of-Retirement-Reporting").Worksheets(1).Range("D2:D316")
    Set resultSegments = Application.Workbooks("Point-of-Retirement-Reporting").Worksheets(1).Range("E2:E316")
    Set resultEventTotals = Application.Workbooks("Point-of-Retirement-Reporting").Worksheets(1).Range("F2:F316")

    ' Main loop
    For counter = 1 To startRange.Count


        If InStr(startRange(counter, 1), "-") > 0 Then
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
        End If
    Next counter
    ' End main loop


    ' Define headings
    Application.Workbooks("Point-of-Retirement-Reporting").Worksheets(1).Range("A1") = "ID"
    Application.Workbooks("Point-of-Retirement-Reporting").Worksheets(1).Range("B1") = "Name"
    Application.Workbooks("Point-of-Retirement-Reporting").Worksheets(1).Range("C1") = "Seller_Code"
    Application.Workbooks("Point-of-Retirement-Reporting").Worksheets(1).Range("D1") = "Event_Label"
    Application.Workbooks("Point-of-Retirement-Reporting").Worksheets(1).Range("E1") = "Segment"
    Application.Workbooks("Point-of-Retirement-Reporting").Worksheets(1).Range("F1") = "Event_Total"

    ' Delete empty rows
    Application.Workbooks("Point-of-Retirement-Reporting").Worksheets(1).Range("A2:F316").Select
    Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete

    Debug.Print "Finished!"
End Sub


'Sub sortResults()

 '   Dim rows As Range

'    Set rows = Application.Workbooks("Point-of-Retirement-Reporting").Worksheets(1).Range("A2:A316")

'    For Each cell In rows
'        For Each innerCell In rows
'            If cell.Value = innerCell.Value Then

'        Next innerCell
'    Next cell


'End Sub




Function parseDataString(inputData As String, delimiter As String) As String()
    Dim all As Variant, returnVal(2) As String

        all = Split(inputData, delimiter)

        returnVal(0) = Trim(all(0))
        returnVal(1) = Trim(all(1))
        returnVal(2) = Trim(all(2))

        parseDataString = returnVal

End Function


